from pathlib import Path
import sqlite3
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Manufacturing Intelligence', page_icon='🏭', layout='wide')
st.title('🏭 Manufacturing Intelligence')
st.caption('Jeeves inventory • purchasing and manufacturing • BAT lot intelligence')
st.warning('Planning support only. Verify inventory, reservations, lot status, locations, and order status in Jeeves/BAT before acting. Century-adjusted dates require review.')
st.sidebar.title('Navigation')
page = st.sidebar.radio('Go to', ['Overview', 'Action Center', 'Inventory', 'Purchasing', 'Manufacturing', 'Hold Monitor', 'Expiration'], label_visibility='collapsed')
st.sidebar.divider()
MASTER_PATH = Path(__file__).resolve().parent / 'master_parts.xlsx'
REVIEW_DB_PATH = Path(__file__).resolve().parent / 'material_reviews.db'



def initialize_review_db():
    with sqlite3.connect(REVIEW_DB_PATH) as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS material_reviews (
            item_no TEXT NOT NULL,
            action_type TEXT NOT NULL,
            status TEXT NOT NULL,
            notes TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            issue_signature TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (item_no, action_type)
        )''')
        columns = {row[1] for row in conn.execute('PRAGMA table_info(material_reviews)')}
        if 'issue_signature' not in columns:
            conn.execute("ALTER TABLE material_reviews ADD COLUMN issue_signature TEXT NOT NULL DEFAULT ''")


def load_review(item_no, action_type):
    with sqlite3.connect(REVIEW_DB_PATH) as conn:
        return conn.execute(
            'SELECT status, notes, issue_signature FROM material_reviews WHERE item_no=? AND action_type=?',
            (item_no, action_type)
        ).fetchone()


def save_review(item_no, action_type, status, notes, issue_signature):
    with sqlite3.connect(REVIEW_DB_PATH) as conn:
        conn.execute('''INSERT INTO material_reviews (item_no, action_type, status, notes, issue_signature)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(item_no, action_type) DO UPDATE SET
                status=excluded.status, notes=excluded.notes, issue_signature=excluded.issue_signature,
                updated_at=CURRENT_TIMESTAMP''',
            (item_no, action_type, status, notes, issue_signature))


def read_excel(source):
    df = pd.read_excel(source, dtype={'Item No.': str, 'Batch/Lot No.': str})
    df.columns = df.columns.astype(str).str.strip()
    if 'Item No.' in df.columns:
        df['Item No.'] = df['Item No.'].astype('string').str.strip()
    return df


def required_columns(df, columns, source):
    missing = [col for col in columns if col not in df.columns]
    if missing:
        st.error(f"{source}: missing columns: {', '.join(missing)}")
        return False
    return True


def show_table(df, columns=None):
    st.dataframe(df if columns is None else df[columns], width='stretch', hide_index=True)


def parse_bat_expiration(value):
    """Parse Excel serial dates, actual datetimes, and BAT date strings."""
    if pd.isna(value) or str(value).strip() == '':
        return pd.NaT, False
    raw = str(value).strip()
    try:
        number = float(raw)
        if number.is_integer() and 1 <= number <= 100000:
            date = pd.to_datetime(number, unit='D', origin='1899-12-30')
            if 1930 <= date.year <= 1999:
                return date + pd.DateOffset(years=100), True
            return date, False
    except (ValueError, TypeError, OverflowError):
        pass
    date = pd.to_datetime(value, errors='coerce')
    if pd.isna(date):
        return pd.NaT, False
    if 1930 <= date.year <= 1999:
        return date + pd.DateOffset(years=100), True
    return date, False


def add_shipdown_fields(df):
    """Compute the fixed six-calendar-month ship-down cutoff for finished goods."""
    df = df.copy()
    fg = df['Inventory Category'].eq('Finished Good')
    df['Ship-Down Date'] = pd.NaT
    df.loc[fg, 'Ship-Down Date'] = df.loc[fg, 'Expiration date'].apply(
        lambda d: d - pd.DateOffset(months=6) if pd.notna(d) else pd.NaT
    )
    days = (pd.to_datetime(df['Ship-Down Date']) - pd.Timestamp.today().normalize()).dt.days
    df['Days to Ship-Down'] = days
    df['Ship-Down Status'] = 'Not applicable (chemical)'
    df.loc[fg, 'Ship-Down Status'] = days[fg].apply(
        lambda n: 'Unknown — verify expiration' if pd.isna(n) else
        'Past ship-down' if n < 0 else
        'Ship-down today' if n == 0 else
        'Within 30 days' if n <= 30 else
        'Within 60 days' if n <= 60 else
        'Within 90 days' if n <= 90 else 'Over 90 days'
    )
    return df


def expiration_review_mask(df):
    """Flag actual expiration risks only."""
    today = pd.Timestamp.today().normalize()

    fg = df['Inventory Category'].eq('Finished Good')
    threshold = today + pd.DateOffset(months=6)
    dates = df['Expiration date']

    return (
        (fg & dates.le(threshold))
        | (~fg & df['Days Until Expiration'].le(90))
        | dates.isna()
    )


def expiry_label(days):
    if pd.isna(days):
        return 'Unknown — review'
    if days < 0:
        return 'Expired'
    if days <= 30:
        return 'Within 30 days'
    if days <= 60:
        return 'Within 60 days'
    if days <= 90:
        return 'Within 90 days'
    return 'Over 90 days'


with st.sidebar.expander('📁 ERP imports', expanded=True):
    inventory_file = st.file_uploader('Jeeves inventory export', type='xlsx', key='inventory')
    bat_file = st.file_uploader('BAT all positive-balance lots export', type='xlsx', key='bat')

master = None
if not MASTER_PATH.exists():
    st.error('master_parts.xlsx was not found beside app.py. Add it to the project folder.')
else:
    try:
        master = read_excel(MASTER_PATH)
        if required_columns(master, ['Item No.', 'Item description'], 'Master list'):
            master = master.dropna(subset=['Item No.']).copy()
            master = master[master['Item No.'].ne('')].copy()
            duplicates = master[master.duplicated('Item No.', keep=False)]
            if not duplicates.empty:
                st.warning(f'Master list contains {len(duplicates)} duplicate rows; keeping first occurrence.')
            master = master.drop_duplicates('Item No.', keep='first')
            st.sidebar.success(f'Master list loaded: {len(master)} unique part numbers')
        else:
            master = None
    except Exception as exc:
        st.error(f'Could not read master_parts.xlsx: {exc}')

bat = None
lot_summary = None
if bat_file is not None:
    try:
        imported = read_excel(bat_file)
        needed = ['Item No.', 'Item description', 'Batch/Lot No.', 'Batch/Lot balance', 'Batch/Lot status', 'Expiration date']
        if required_columns(imported, needed, 'BAT'):
            imported['Batch/Lot balance'] = pd.to_numeric(imported['Batch/Lot balance'], errors='coerce')
            invalid_balances = imported['Batch/Lot balance'].isna().sum()
            if invalid_balances:
                st.warning(f'BAT contains {invalid_balances} rows with invalid lot balances; excluded from lot calculations.')
            bat = imported[imported['Batch/Lot balance'] > 0].copy()
            bat = bat.dropna(subset=['Item No.'])
            bat = bat[bat['Item No.'].ne('')].copy()
            bat['Normalized Status'] = bat['Batch/Lot status'].astype('string').str.strip().str.casefold().fillna('')
            bat['Lot Classification'] = bat['Normalized Status'].map({'hold': 'Hold', 'released': 'Released'}).fillna('Other / review')
            bat['Held Qty'] = bat['Batch/Lot balance'].where(bat['Lot Classification'].eq('Hold'), 0)
            bat['Released Qty'] = bat['Batch/Lot balance'].where(bat['Lot Classification'].eq('Released'), 0)
            bat['Other Status Qty'] = bat['Batch/Lot balance'].where(bat['Lot Classification'].eq('Other / review'), 0)
            lot_summary = bat.groupby('Item No.', as_index=False).agg(
                BAT_Total=('Batch/Lot balance', 'sum'),
                BAT_Held=('Held Qty', 'sum'),
                BAT_Released=('Released Qty', 'sum'),
                BAT_Other=('Other Status Qty', 'sum'),
                BAT_Lots=('Batch/Lot No.', 'size'),
            )
            st.sidebar.success(f'BAT loaded: {len(bat)} positive-balance lots across {len(lot_summary)} items')
    except Exception as exc:
        st.error(f'BAT import or processing error: {exc}')

# Prepare expiration alerts before rendering the Daily Action Queue.
# Use positive-balance BAT lots; don't assume held or expired lots are usable.
expiry_actions = pd.DataFrame()
if bat is not None and master is not None:
    expiry_actions = bat.loc[
        bat['Item No.'].str.startswith(('701', '501'), na=False)
        & bat['Item No.'].isin(master['Item No.'])
    ].copy()
    parsed = expiry_actions['Expiration date'].apply(parse_bat_expiration)
    expiry_actions['Expiration date'] = pd.to_datetime(
        parsed.map(lambda pair: pair[0]), errors='coerce')
    expiry_actions['Date Corrected'] = parsed.map(lambda pair: pair[1])
    expiry_actions['Days Until Expiration'] = (
        expiry_actions['Expiration date'] - pd.Timestamp.today().normalize()).dt.days
    expiry_actions['Inventory Category'] = expiry_actions['Item No.'].apply(
        lambda pn: 'Finished Good' if str(pn).startswith('501') else 'Chemical')
    expiry_actions['Expiration Status'] = expiry_actions['Days Until Expiration'].apply(expiry_label)
    expiry_actions = add_shipdown_fields(expiry_actions)

if inventory_file is None and page not in ('Hold Monitor', 'Expiration'):
    st.info('Upload a Jeeves inventory export in the sidebar to begin.')
if bat_file is None and page in ('Hold Monitor', 'Expiration'):
    st.info('Upload a BAT export in the sidebar to begin.')

if inventory_file is not None:
    try:
        inv = read_excel(inventory_file)
        needed = ['Item No.', 'Item description', 'Inventory balance', 'RsrvQty', 'OrdPoint', 'OrdQtyPO', 'OrdQtyMO']
        if required_columns(inv, needed, 'Jeeves'):
            inv = inv.dropna(subset=['Item No.']).copy()
            inv = inv[inv['Item No.'].ne('')].copy()
            inv['Item Type'] = 'Other / review'
            inv.loc[inv['Item No.'].str.startswith('501', na=False), 'Item Type'] = 'Finished good (MO)'
            inv.loc[inv['Item No.'].eq('503-347'), 'Item Type'] = 'Stocked subassembly (MO)'
            if master is not None:
                inv['On Master List'] = inv['Item No.'].isin(master['Item No.'])
                missing = master.loc[~master['Item No.'].isin(inv['Item No.']), ['Item No.', 'Item description']]
                if not missing.empty:
                    st.warning(f'{len(missing)} master-list items are missing from the Jeeves export.')
                    with st.expander('⚠️ Missing monitored part numbers'):
                        show_table(missing)
            else:
                inv['On Master List'] = False

            st.sidebar.subheader('🔎 Inventory filters')
            scope = st.sidebar.radio('View', ['Monitored parts', 'All exported items'], horizontal=True,
                             index=0 if master is not None else 1, disabled=master is None)
            view = inv[inv['On Master List']].copy() if scope == 'Monitored parts' else inv.copy()
            query = st.sidebar.text_input('Search item number or description', placeholder='Type part number or name')
            if query:
                view = view[view['Item No.'].str.contains(query, case=False, regex=False, na=False) |
                            view['Item description'].astype(str).str.contains(query, case=False, regex=False, na=False)]
            if master is not None and scope == 'Monitored parts':
                st.caption('Showing only part numbers in master_parts.xlsx. Choose All exported items to view others.')

            nums = ['Inventory balance', 'RsrvQty', 'OrdPoint', 'OrdQtyPO', 'OrdQtyMO']
            for col in nums:
                view[col] = pd.to_numeric(view[col], errors='coerce')
            invalid = view[view[['Inventory balance', 'RsrvQty']].isna().any(axis=1)]
            if not invalid.empty:
                st.warning(f'{len(invalid)} displayed inventory rows have missing/invalid stock or reservations; excluded.')
            view = view.dropna(subset=['Inventory balance', 'RsrvQty']).copy()
            view['Net Stock'] = view['Inventory balance'] - view['RsrvQty']
            view['Jeeves Reorder Alert'] = view['OrdPoint'].gt(0) & view['Net Stock'].lt(view['OrdPoint'])
            view['Jeeves Shortage Alert'] = view['Net Stock'] < 0
            view['Order Qty for Review'] = view['OrdQtyMO'].where(
                view['Item Type'].isin(['Finished good (MO)', 'Stocked subassembly (MO)']), view['OrdQtyPO'])

            # Only reconcile item-level BAT totals against a single item-level Jeeves row.
            # Multi-location items require location-aware reconciliation, not a blind merge.
            view['Jeeves Item Rows'] = view.groupby('Item No.')['Item No.'].transform('size')
            if lot_summary is not None:
                view = view.merge(lot_summary, on='Item No.', how='left', validate='many_to_one')
            else:
                for col in ['BAT_Total', 'BAT_Held', 'BAT_Released', 'BAT_Other', 'BAT_Lots']:
                    view[col] = float('nan')
            view['BAT Difference'] = view['Inventory balance'] - view['BAT_Total']
            view['BAT Reconciled'] = (view['Jeeves Item Rows'].eq(1) & view['BAT_Total'].notna() &
                                      view['BAT Difference'].abs().le(0.01))
            view['Adjusted Net Stock'] = (view['BAT_Released'] - view['RsrvQty']).where(view['BAT Reconciled'])
            view['Adjusted Shortage Alert'] = view['Adjusted Net Stock'].lt(0) & view['BAT Reconciled']
            view['Adjusted Reorder Alert'] = (view['OrdPoint'].gt(0) & view['Adjusted Net Stock'].lt(view['OrdPoint']) & view['BAT Reconciled'])
            view['Hidden Shortage'] = view['Adjusted Shortage Alert'] & ~view['Jeeves Shortage Alert']
            view['Adjusted Shortage Qty'] = (-view['Adjusted Net Stock']).clip(lower=0)
            view['Review Reason'] = 'Ready — BAT reconciled'
            view.loc[view['BAT_Total'].isna(), 'Review Reason'] = 'No BAT lots — cannot verify availability'
            view.loc[view['BAT_Total'].notna() & ~view['BAT Reconciled'], 'Review Reason'] = 'BAT/Jeeves mismatch or multiple inventory rows — review'
            view.loc[view['BAT Reconciled'] & view['BAT_Other'].gt(0), 'Review Reason'] = 'Other lot statuses excluded from usable stock — review'

            if page == "Inventory":
                st.header('📦 Inventory Overview')
                a, b, c, d = st.columns(4)
                a.metric('Inventory rows', len(view))
                b.metric('Jeeves shortages', int(view['Jeeves Shortage Alert'].sum()))
                c.metric('Hold-adjusted shortages', int(view['Adjusted Shortage Alert'].sum()))
                d.metric('Hidden shortages', int(view['Hidden Shortage'].sum()))
                st.caption('Adjusted alerts are shown only for items whose positive-balance BAT lot totals reconcile to Jeeves. Missing or mismatched items require manual review.')
                cols = ['Item No.', 'Item description', 'Item Type', 'Inventory balance', 'RsrvQty', 'Net Stock',
                        'BAT_Held', 'BAT_Released', 'BAT_Other', 'Adjusted Net Stock', 'OrdPoint',
                        'Jeeves Shortage Alert', 'Adjusted Shortage Alert', 'Hidden Shortage', 'Review Reason']
                show_table(view, cols)
                st.subheader('🚨 Shortages and hidden shortages')
                show_table(view[view['Jeeves Shortage Alert'] | view['Adjusted Shortage Alert']], cols)
                if lot_summary is not None:
                    needs_review = view[~view['BAT Reconciled'] | view['BAT_Other'].gt(0)]
                    with st.expander(f'⚠️ BAT reconciliation / status review ({len(needs_review)} rows)'):
                        show_table(needs_review, ['Item No.', 'Item description', 'Inventory balance', 'BAT_Total',
                                                 'BAT Difference', 'BAT_Held', 'BAT_Released', 'BAT_Other', 'Review Reason'])

            # A missing/zero reorder point disables STOCK replenishment, not customer-demand scheduling.
            # EOQ is optional: no EOQ column was present in the required export schema.
            eoq_candidates = ['EOQ', 'Eoq', 'Economic order quantity', 'Economic Order Quantity']
            eoq_col = next((col for col in eoq_candidates if col in inv.columns), None)
            if eoq_col:
                # Map directly rather than merging: if the export already contains
                # an EOQ column, merging would rename it EOQ_x / EOQ_y.
                eoq_lookup = (inv[['Item No.', eoq_col]]
                              .drop_duplicates('Item No.')
                              .set_index('Item No.')[eoq_col])
                view['EOQ'] = pd.to_numeric(view['Item No.'].map(eoq_lookup), errors='coerce')
            else:
                view['EOQ'] = float('nan')

            mo_types = ['Finished good (MO)', 'Stocked subassembly (MO)']
            manufactured = view[view['Item Type'].isin(mo_types)].copy()
            manufactured['Planning Usable Stock'] = manufactured['BAT_Released'].where(
                manufactured['BAT Reconciled'], manufactured['Inventory balance'])
            manufactured['Demand Shortage Qty'] = (
                manufactured['RsrvQty'] - manufactured['Planning Usable Stock']).clip(lower=0)
            manufactured['Additional MO Qty to Review'] = (
                manufactured['Demand Shortage Qty'] - manufactured['OrdQtyMO'].fillna(0)).clip(lower=0)
            manufactured['Reorder Trigger'] = (manufactured['Jeeves Reorder Alert'] |
                                                 manufactured['Adjusted Reorder Alert'])
            manufactured['Demand Trigger'] = manufactured['Demand Shortage Qty'].gt(0)
            manufactured = manufactured[manufactured['Reorder Trigger'] | manufactured['Demand Trigger']].copy()

            def manufacturing_action(row):
                if row['Demand Trigger']:
                    if pd.isna(row['OrdQtyMO']):
                        return 'Customer demand — MO quantity missing; verify Jeeves'
                    if row['Additional MO Qty to Review'] > 0:
                        return 'Customer demand — review additional MO scheduling'
                    return 'Customer demand quantity covered by MO — verify due dates'
                if row['Reorder Trigger']:
                    if pd.notna(row['EOQ']) and row['EOQ'] > 0:
                        return 'Stock replenishment — review EOQ build and existing MOs'
                    return 'Stock replenishment — EOQ missing; review build quantity'
                return 'No action'

            manufactured['Manufacturing Action'] = manufactured.apply(manufacturing_action, axis=1)
            manufactured['EOQ Build Qty for Review'] = manufactured['EOQ'].where(
                manufactured['Reorder Trigger'] & manufactured['EOQ'].gt(0))
            if page == "Manufacturing":
                st.header('🏭 Manufacturing Action Center')
                st.caption('Customer reservations trigger OEM scheduling even with a zero/blank reorder point. '
                           'MO coverage is quantity-only: verify order allocation, due dates and MO status in Jeeves. '
                           'For BAT mismatches, usable stock falls back to Jeeves and requires review.')
                show_table(manufactured, ['Item No.', 'Item description', 'Inventory balance', 'RsrvQty',
                                          'BAT_Held', 'Adjusted Net Stock', 'OrdPoint', 'EOQ',
                                          'OrdQtyMO', 'Demand Shortage Qty', 'Additional MO Qty to Review',
                                          'EOQ Build Qty for Review', 'Hidden Shortage', 'Review Reason',
                                          'Manufacturing Action'])
                if eoq_col is None:
                    st.info('EOQ was not found in this Jeeves export. Add its exact column to the export '
                            'to populate EOQ build recommendations; OEM demand scheduling works without it.')

            purchasing = view[(view['Item Type'] == 'Other / review') &
                              (view['Jeeves Shortage Alert'] | view['Adjusted Shortage Alert'] |
                               view['Jeeves Reorder Alert'] | view['Adjusted Reorder Alert'])].copy()
            if page == "Purchasing":
                st.header('🛒 Purchasing Action Center')
                st.caption('Purchased materials appear for shortages OR stock below reorder point. Verify sourcing, existing POs, and demand before ordering.')
            purchasing['Shortage Qty'] = (-purchasing['Net Stock']).clip(lower=0)
            purchasing['Planning Shortage Qty'] = purchasing['Adjusted Shortage Qty'].where(
                purchasing['BAT Reconciled'], purchasing['Shortage Qty'])
            purchasing['Planning Net Stock'] = purchasing['Adjusted Net Stock'].where(
                purchasing['BAT Reconciled'], purchasing['Net Stock'])
            purchasing['Reorder Gap Qty'] = (
                purchasing['OrdPoint'] - purchasing['Planning Net Stock']).clip(lower=0)
            purchasing['Required Qty for Review'] = purchasing[
                ['Planning Shortage Qty', 'Reorder Gap Qty']].max(axis=1)
            purchasing['Reorder Reason'] = purchasing.apply(
                lambda r: 'Below reorder point' if r['Reorder Gap Qty'] > 0
                else 'Stock shortage', axis=1)
            def po_action(row):
                qty = row['OrdQtyPO']
                if pd.isna(qty):
                    return 'PO quantity missing — verify'
                if qty <= 0:
                    return 'Below reorder point — review purchasing' if row['Reorder Gap Qty'] > 0 else 'No PO recorded'
                if qty < row['Required Qty for Review']:
                    return 'PO quantity below planning gap — review'
                return 'PO recorded — verify sent / delivery'
            purchasing['Purchasing Action'] = purchasing.apply(po_action, axis=1)
            purchasing['Uncovered Qty if PO Arrives'] = (purchasing['Required Qty for Review'] - purchasing['OrdQtyPO']).clip(lower=0)
            if page == "Purchasing":
                show_table(purchasing, ['Item No.', 'Item description', 'BAT_Held', 'Net Stock', 'Adjusted Net Stock',
                                        'OrdPoint', 'Reorder Reason', 'Reorder Gap Qty', 'Planning Shortage Qty',
                                        'Required Qty for Review', 'OrdQtyPO', 'Uncovered Qty if PO Arrives',
                                        'Review Reason', 'Purchasing Action'])
                st.caption('Reorder Gap Qty is a planning signal, not an automatic order quantity or EOQ. Existing POs are considered for the uncovered gap, but PO entries are not proof of supplier acknowledgment.')

            if page == "Action Center":
                st.header('📋 Daily Action Queue')
                st.caption('Start with Needs Attention. Saved reviews are local planning notes, not changes to Jeeves.')
            try:
                initialize_review_db()
                action_frames = []
                for action_type, source_df, action_col in [
                    ('Purchasing', purchasing, 'Purchasing Action'),
                    ('Manufacturing', manufactured, 'Manufacturing Action'),
                ]:
                    for _, record in source_df.iterrows():
                        item_no = str(record['Item No.'])
                        # Capture relevant current conditions to detect a changed issue after resolution.
                        signature_cols = ['Inventory balance', 'RsrvQty', 'OrdPoint',
                                          'OrdQtyPO', 'OrdQtyMO', 'BAT_Held', 'BAT_Released',
                                          'BAT Reconciled', action_col]
                        signature = '|'.join(str(record.get(col, '')) for col in signature_cols)
                        previous = load_review(item_no, action_type)
                        saved_status = previous[0] if previous else 'Needs review'
                        changed = bool(previous and saved_status == 'Resolved' and
                                       previous[2] != signature)
                        effective_status = 'Needs review' if changed else saved_status
                        group = ('Resolved' if effective_status == 'Resolved' else
                                 'In Progress' if effective_status in ['Waiting on PO', 'Waiting on supplier'] else
                                 'Needs Attention')
                        action_frames.append({
                            'Item No.': item_no,
                            'Item description': record['Item description'],
                            'Area': action_type,
                            'Review status': effective_status,
                            'Current recommendation': record[action_col],
                            'Queue': group,
                            'Changed since resolved': changed,
                            '_signature': signature,
                        })
                # One review per part number for expiration risk, with lot-level detail.
                # A part can have several lots, including healthy alternatives.
                if not expiry_actions.empty:
                    fg_shipdown_soon = (
                        expiry_actions['Inventory Category'].eq('Finished Good')
                        & expiry_actions['Days to Ship-Down'].between(0, 90)
                    )
                    at_risk = expiry_actions[
                        expiration_review_mask(expiry_actions) | fg_shipdown_soon
                    ].copy()
                    for item_no, lots in at_risk.groupby('Item No.', sort=False):
                        all_lots = expiry_actions[expiry_actions['Item No.'].eq(item_no)]
                        category = lots['Inventory Category'].iloc[0]
                        area = 'Expiration — ' + category
                        # Signature includes all lots, statuses, balances, and dates.
                        # It changes if a lot expires, a lot is added, or stock changes.
                        signature = '|'.join(sorted(
                            f"{r['Batch/Lot No.']}:{r['Batch/Lot balance']}:"
                            f"{r['Lot Classification']}:{r['Expiration date']}:"
                            f"{r['Expiration Status']}:{r['Ship-Down Status']}:{r['Date Corrected']}"
                            for _, r in all_lots.iterrows()
                        ))
                        previous = load_review(str(item_no), area)
                        saved_status = previous[0] if previous else 'Needs review'
                        changed = bool(previous and saved_status == 'Resolved'
                                       and previous[2] != signature)
                        effective_status = 'Needs review' if changed else saved_status
                        group = ('Resolved' if effective_status == 'Resolved' else
                                 'In Progress' if effective_status in
                                 ['Waiting on PO', 'Waiting on supplier'] else 'Needs Attention')
                        risky_qty = lots['Batch/Lot balance'].sum()
                        healthy = all_lots[
                            all_lots['Lot Classification'].eq('Released')
                            & ~expiration_review_mask(all_lots)
                            & ~all_lots['Date Corrected']
                        ]
                        healthy_qty = healthy['Batch/Lot balance'].sum()
                        unknown = lots['Days Until Expiration'].isna().sum()
                        shipdown_note = ''
                        if category == 'Finished Good':
                            past = int(lots['Ship-Down Status'].isin(['Past ship-down', 'Ship-down today']).sum())
                            soon = int(lots['Days to Ship-Down'].between(1, 90).sum())
                            shipdown_note = f'{past} lot(s) at/past ship-down; {soon} approaching within 90 days. '
                        recommendation = (
                            shipdown_note +
                            f"{len(lots)} lot(s) need expiry review; {risky_qty:g} units flagged; "
                            f"{healthy_qty:g} units in released lots beyond the review window; "
                            f"{unknown} unknown date(s). Check reservations, demand, "
                            + ('other usable lots and MO requirements in Jeeves.'
                               if category == 'Finished Good'
                               else 'other usable lots and purchasing requirements in Jeeves.')
                        )
                        action_frames.append({
                            'Item No.': str(item_no),
                            'Item description': lots['Item description'].iloc[0],
                            'Area': area,
                            'Review status': effective_status,
                            'Current recommendation': recommendation,
                            'Queue': group,
                            'Changed since resolved': changed,
                            '_signature': signature,
                        })
                if not action_frames:
                    if page in ("Overview", "Action Center"):
                        st.header('📊 Executive Overview')
                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric('Monitored inventory rows', len(view))
                        k2.metric('Stock shortages', int(
                            (view['Jeeves Shortage Alert'] | view['Adjusted Shortage Alert']).sum()))
                        k3.metric('Held BAT lots', int(
                            bat['Lot Classification'].eq('Hold').sum()) if bat is not None else 0)
                        k4.metric('Expiration-risk lots', int(
                            (expiration_review_mask(expiry_actions) |
                             (expiry_actions['Inventory Category'].eq('Finished Good') &
                              expiry_actions['Days to Ship-Down'].between(0, 90))).sum())
                            if not expiry_actions.empty else 0)
                        st.success('No currently flagged action items.')
                else:
                    queue = pd.DataFrame(action_frames)
                    # Collapse duplicate item/area rows in the selector; review is item-level.
                    queue = queue.drop_duplicates(['Item No.', 'Area'], keep='first')
                    # Executive overview: use the same filtered planning data and action queue.
                    if page == "Overview":
                        st.header('📊 Executive Overview')
                        st.caption('Live snapshot of the selected Jeeves inventory scope and BAT imports. '
                                   'Counts reflect currently flagged actions, not all historical reviews.')

                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric('Monitored inventory rows', len(view))
                        k2.metric('Stock shortages', int(
                            (view['Jeeves Shortage Alert'] | view['Adjusted Shortage Alert']).sum()))
                        k3.metric('Held BAT lots', int(
                            bat['Lot Classification'].eq('Hold').sum()) if bat is not None else 0)
                        k4.metric('Expiration-risk lots', int(
                            (expiration_review_mask(expiry_actions) |
                             (expiry_actions['Inventory Category'].eq('Finished Good') &
                              expiry_actions['Days to Ship-Down'].between(0, 90))).sum())
                            if not expiry_actions.empty else 0)

                        q1, q2, q3 = st.columns(3)
                        q1.metric('Needs Attention', int(queue['Queue'].eq('Needs Attention').sum()))
                        q2.metric('In Progress', int(queue['Queue'].eq('In Progress').sum()))
                        q3.metric('Resolved', int(queue['Queue'].eq('Resolved').sum()))

                        chart_left, chart_right = st.columns(2)
                        with chart_left:
                            st.subheader('Action items by area')
                            area_counts = queue.groupby('Area').size().rename('Action items')
                            st.bar_chart(area_counts)
                        with chart_right:
                            st.subheader('Action queue status')
                            status_counts = queue['Queue'].value_counts().reindex(
                                ['Needs Attention', 'In Progress', 'Resolved'], fill_value=0)
                            st.bar_chart(status_counts.rename('Action items'))

                        with st.expander('🔍 Executive overview definitions'):
                            st.caption('Stock shortages combine Jeeves and reconciled hold-adjusted alerts. '
                                       'Held BAT lots count all imported positive-balance BAT lots, '
                                       'regardless of the selected inventory filter. '
                                       'Expiration-risk lots cover master-listed finished goods within six '
                                       'months of expiration, chemicals within 90 days, unknown dates, '
                                       'and finished goods approaching ship-down within 90 days. '
                                       'Action items are unique per part number and action area. '
                                       'This overview is informational; confirm actions in Jeeves/BAT.')
                        st.divider()

                    if page == "Action Center":
                        groups = ['Needs Attention', 'In Progress', 'Resolved']
                        tabs = st.tabs([f'{name} ({int(queue["Queue"].eq(name).sum())})' for name in groups])
                        for tab, group_name in zip(tabs, groups):
                            with tab:
                                subset = queue[queue['Queue'].eq(group_name)]
                                if subset.empty:
                                    st.info('No items in this queue.')
                                else:
                                    show_table(subset, ['Item No.', 'Item description', 'Area',
                                                        'Review status', 'Current recommendation',
                                                        'Changed since resolved'])
                        st.subheader('📝 Review an action item')
                        st.caption('Select a queue below, then choose the material to update.')
                        chosen_group = st.selectbox('Review queue', groups, key='review_queue_filter')
                        selectable = queue[queue['Queue'].eq(chosen_group)]
                        if selectable.empty:
                            st.info('Choose a queue containing items to review.')
                        else:
                            options = {
                                f"{r['Area']} | {r['Item No.']} | {r['Item description']}": r
                                for _, r in selectable.iterrows()
                            }
                            selection = st.selectbox('Choose an action item', list(options), key='review_item')
                            selected = options[selection]
                            item_no, action_type = selected['Item No.'], selected['Area']
                            previous = load_review(item_no, action_type)
                            statuses = ['Needs review', 'Waiting on PO', 'Waiting on supplier', 'Resolved']
                            initial_status = selected['Review status']
                            if selected['Changed since resolved']:
                                st.warning('Inventory or order information changed since this item was resolved. Review again.')
                            with st.form(key=f'review_form_{action_type}_{item_no}'):
                                status = st.selectbox('Review status', statuses,
                                                      index=statuses.index(initial_status))
                                notes = st.text_area('Review notes', value=previous[1] if previous else '',
                                                     placeholder='What did you check? What needs to happen next?')
                                submitted = st.form_submit_button('Save review')
                            if submitted:
                                save_review(item_no, action_type, status, notes, selected['_signature'])
                                st.success(f'Saved review for {item_no} ({action_type}).')
                                st.rerun()
                        st.caption('Only currently flagged items appear in these queues. Expiration items are grouped by part number; inspect lot-level details in the Inventory Expiration Monitor. A resolved item is shown '
                                   'for re-review when its tracked inventory/order values change. '
                                   'Review statuses are saved in material_reviews.db.')
            except sqlite3.Error as exc:
                st.error(f'Could not access local review database: {exc}')

    except Exception as exc:
        st.error(f'Jeeves import or processing error: {exc}')

if bat is not None:
    if page == "Hold Monitor":
        st.header('🧩 BAT Lot Intelligence')
        st.subheader('🛑 Inventory Hold Monitor')
        hold_scope = st.radio(
            'Hold monitor scope',
            ['Monitored parts', 'All BAT items'],
            index=0 if master is not None else 1,
            disabled=master is None,
            horizontal=True,
            key='hold_scope',
        )
        hold_bat = bat[bat['Item No.'].isin(master['Item No.'])].copy() if hold_scope == 'Monitored parts' else bat.copy()
        held_lots = hold_bat[hold_bat['Lot Classification'].eq('Hold')].copy()
        other_lots = hold_bat[hold_bat['Lot Classification'].eq('Other / review')].copy()
        a, b, c = st.columns(3)
        a.metric('Lots on hold', len(held_lots))
        b.metric('Items with held lots', held_lots['Item No.'].nunique())
        c.metric('Lots with other statuses', len(other_lots))
        st.caption('Hold monitor counts and rows follow the selected scope. All BAT lots remain available for inventory reconciliation and expiration analysis.')
        hold_cols = ['Item No.', 'Item description', 'Batch/Lot No.', 'Batch/Lot status', 'Batch/Lot balance']
        show_table(held_lots, hold_cols)
        with st.expander('🔎 Other BAT lot statuses — review before treating as usable'):
            show_table(other_lots, hold_cols)

    if page == "Expiration":
        # Monitor master-listed 701 chemicals and 501 finished goods from positive-balance BAT lots.
        st.subheader('📅 Inventory Expiration Monitor — Chemicals & Finished Goods')
        if master is None:
            st.error('Inventory Expiration Monitor requires a valid master_parts.xlsx file.')
            expiring_inventory = bat.iloc[0:0].copy()
        else:
            eligible = bat['Item No.'].str.startswith(('701', '501'), na=False)
            monitored = bat['Item No.'].isin(master['Item No.'])
            expiring_inventory = bat.loc[eligible & monitored].copy()
            st.caption('Showing master-listed 701-series chemicals and 501-series finished goods with positive BAT lot balances.')

        expiring_inventory['Inventory Category'] = 'Chemical'
        expiring_inventory.loc[
            expiring_inventory['Item No.'].str.startswith('501', na=False), 'Inventory Category'
        ] = 'Finished Good'

        expiring_inventory['Original Expiration Date'] = expiring_inventory['Expiration date'].astype(str)
        parsed_dates = expiring_inventory['Expiration date'].apply(parse_bat_expiration)
        expiring_inventory['Expiration date'] = pd.to_datetime(
            parsed_dates.map(lambda result: result[0]), errors='coerce'
        )
        expiring_inventory['Date Corrected'] = parsed_dates.map(lambda result: result[1])
        expiring_inventory['Days Until Expiration'] = (
            expiring_inventory['Expiration date'] - pd.Timestamp.today().normalize()
        ).dt.days
        expiring_inventory['Expiration Status'] = expiring_inventory['Days Until Expiration'].apply(expiry_label)
        expiring_inventory = add_shipdown_fields(expiring_inventory)

        category_filter = st.selectbox(
            'Inventory category', ['All', 'Chemicals', 'Finished Goods'], key='expiration_category'
        )
        expiry_view = expiring_inventory
        if category_filter == 'Chemicals':
            expiry_view = expiring_inventory[expiring_inventory['Inventory Category'].eq('Chemical')]
        elif category_filter == 'Finished Goods':
            expiry_view = expiring_inventory[expiring_inventory['Inventory Category'].eq('Finished Good')]

        st.caption('Century-adjusted dates are provisional; verify them against BAT. Held lots remain visible but are not assumed usable.')
        a, b, c, d = st.columns(4)
        a.metric('Tracked lots', len(expiry_view))
        b.metric('Expired', int(expiry_view['Days Until Expiration'].lt(0).sum()))
        c.metric('Expiring within 30 days', int(expiry_view['Days Until Expiration'].between(0, 30).sum()))
        d.metric('Within review window', int((expiration_review_mask(expiry_view) & expiry_view['Days Until Expiration'].ge(0)).sum()))
        exp_cols = ['Inventory Category', 'Item No.', 'Item description', 'Batch/Lot No.',
                    'Batch/Lot status', 'Batch/Lot balance', 'Original Expiration Date',
                    'Expiration date', 'Date Corrected', 'Days Until Expiration', 'Expiration Status',
                    'Ship-Down Date', 'Days to Ship-Down', 'Ship-Down Status']
        show_table(expiry_view, exp_cols)
        st.subheader('⚠️ Expiration alerts (finished goods: 6 months; chemicals: 90 days; unknown dates)')
        shipdown_soon = (expiry_view['Inventory Category'].eq('Finished Good')
                         & expiry_view['Days to Ship-Down'].between(0, 90))
        show_table(expiry_view[expiration_review_mask(expiry_view) | shipdown_soon], exp_cols)
        st.caption('Ship-down is six calendar months before expiration for finished goods. The 90-day upcoming cutoff alerts extend visibility beyond the six-month expiration window. Confirm shipment eligibility in Jeeves and customer requirements.')
        with st.expander('🔎 Date validation — century-adjusted lots'):
            show_table(expiry_view[expiry_view['Date Corrected']], exp_cols)

st.divider()
st.caption('Manufacturing Intelligence • Local development version')

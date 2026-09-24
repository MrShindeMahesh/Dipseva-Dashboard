import store, summary
rows = store.read_ledger()
items, custs = store.read_masters()
s = summary.summarize(rows, items, custs, '2026-07-15')
print('rows', len(rows))
print('billed', s['billed'], 'received', s['received'], 'pending', s['pending'])
yg = [e for e in s['khata'] if e['name'] == 'Yogesh Caterers']
print('Yogesh:', yg)
print('due', s['due_count'], 'low', len(s['low']))
print('top5', [(e['name'], e['pending']) for e in s['khata'][:5]])
st = {e['id']: e for e in s['stock']}
for k in ['CH', 'CTRNG', 'DSH', 'TBL']:
    print(k, st.get(k))

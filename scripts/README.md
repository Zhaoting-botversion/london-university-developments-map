# Price sync script

Use `update_index_prices.py` after editing `University community developments - enriched prices.xlsx`.

From the Obsidian vault root:

```powershell
& 'C:\Users\patri\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '💼 工作专项/London University Developments Map/scripts/update_index_prices.py'
```

If it says `PRICE_LOOKUP would be updated`, apply the update:

```powershell
& 'C:\Users\patri\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' '💼 工作专项/London University Developments Map/scripts/update_index_prices.py' --write
```

The script only replaces the `PRICE_LOOKUP` block in `index.html`. It does not edit map coordinates, commute rows, or `UNIVERSITIES_STUB`.


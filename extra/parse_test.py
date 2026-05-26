#%% header
import pandas as pd
from reader import _get_wb
from reader import _parse_dashboard_MM,_parse_dashboard_arb, _parse_dashboard_MM_sector 

from reader import _parse_option_dashboard, _parse_live_orders, _parse_position, _parse_daily
                   


#%% get wb
wb = _get_wb()
print(f"워크북: {wb.name}")


#%% ── dashboard ──────────────────────────────────────────────────
df = _parse_dashboard_MM(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()

#%% ── dashboard ──────────────────────────────────────────────────
df = _parse_dashboard_arb(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()


#%% ── dashboard ──────────────────────────────────────────────────
df = _parse_dashboard_MM_sector(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()



#%% ── option_dashboard ───────────────────────────────────────────
df = _parse_option_dashboard(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()


#%% ── live_orders ────────────────────────────────────────────────
df = _parse_live_orders(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()


#%% ── position ───────────────────────────────────────────────────
df = _parse_position(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()


#%% ── daily ──────────────────────────────────────────────────────
df = _parse_daily(wb)
print(f"shape   : {df.shape}")
print(f"columns : {df.columns.tolist()}")
print(f"dtypes  :\n{df.dtypes}")
print(f"nulls   :\n{df.isnull().sum()}")
df.head()

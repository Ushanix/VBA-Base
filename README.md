# VBA-Base

DocumentBase, TimeBase, and future Excel/VBA projects share a common set of reusable modules. VBA-Base is the single source of truth for these modules.

## Structure

```
src/
  common/     Shared utility modules (Utl_*)
  data/       Tbl: marker search and DB-like query modules (planned)
  io/         Import/export modules (planned)
templates/    Project-specific templates (Mod_Constants, etc.)
scripts/      Import/export automation scripts
```

## Modules

| ID | Module | Description |
|----|--------|-------------|
| UTL-001 | Utl_File | File operations using FileSystemObject |
| UTL-002 | Utl_Logger | Logging to Immediate Window |
| UTL-003 | Utl_Sheet | Sheet operations (filter, sort, copy) |
| UTL-004 | Utl_Table | Table operations (Tbl marker search, read/write) |
| UTL-005 | Utl_Yaml | Minimal YAML serializer/parser for DataIO |

See [registry.yaml](registry.yaml) for the full module catalog.

## Usage

1. Copy modules from `src/common/` into your project's `excel/src/vba/common/`
2. Copy `templates/Mod_Constants_Template.bas` and customize for your project
3. Follow naming conventions defined in [InformationManagementGuidelines](https://github.com/Ushanix/InformationManagementGuidelines)

## Naming Conventions

All modules follow the prefix-based naming system:

| Prefix | Layer | Purpose |
|--------|-------|---------|
| `Utl_` | Utility | Generic reusable functions |
| `Pst_` | Presentation | UI and event handlers |
| `Mgr_` | Manager | Workflow orchestration |
| `Cfg_` | Config | Constants and configuration |
| `Ent_` | Entry | Tool launch and initialization |
| `Dbg_` | Debug | Testing and validation |
| `Fct_` | Factory | Instance generation |

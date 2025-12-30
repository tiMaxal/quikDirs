# QuikDirs – Help

Voded [vibe-coded] copliot20251230timaxal

Windows 11 does not support alphabetical sorting of pinned Quick Access items.
QuikDirs works around this by pinning a single hub folder and letting Explorer
alphabetically sort the contents of that folder.

QuikDirs provides a stable way to organise folders under a single
Quick Access entry in Windows 11.

## Key points:

- **Hub folder:** C:\$quikDirs
- **Uses NTFS junctions** (not shortcuts) for performance
- **Deleting a junction is safe**
- **Collision naming:** Name[Parent]

## Usage:

1. Right-click a folder → Send To → QuikDirs
2. Junction appears instantly in the hub folder
3. Optional import of pinned Quick Access folders
4. Undo last operations (atomic, multi-folder)

## Safety:

- Only pinned Quick Access items are affected
- Frequent / recent folders are ignored
- No registry hacks are performed

## GUI:

- **Import**: Green button, creates junctions for selected folders
- **Undo Last**: Yellow button, reverses last import operation
- **Exit**: Red button, closes app
- **Close on completion**: Checkbox, optional auto-close after import
- **Help**: Opens this readme
- Scrollable list ensures all pinned Quick Access folders are visible

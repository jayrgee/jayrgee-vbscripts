@echo off

set _target_dir=%appdata%\Microsoft\Windows\SendTo

if exist %_target_dir% (
    copy .\SendTo\ExportVBA.vbs %_target_dir%
    copy .\SendTo\Backup.vbs %_target_dir%
    dir %_target_dir%
)
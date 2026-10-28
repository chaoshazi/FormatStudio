# Vendor directory

Place a LibreOffice Portable installation here as `vendor/libreoffice/` before building the installer.

Expected executable: `vendor/libreoffice/program/soffice.exe`.

The build script records the LibreOffice version and SHA256 in `vendor/manifest.json`.
WPS Office is not bundled; it is detected from the user's system at runtime.
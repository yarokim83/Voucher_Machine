# Version Synchronization Rule

Whenever the project version is updated (e.g. bumping version.txt or tagging a release), you MUST automatically update all 3 locations:
1. version.txt: Set exact version string (e.g. v5.4.0)
2. CHANGELOG.md: Add a release log section for the new version
3. app_gui.py: Update the header badge widget label text (ver_b text="v5.4.0") so the UI header badge ALWAYS displays the exact updated version!

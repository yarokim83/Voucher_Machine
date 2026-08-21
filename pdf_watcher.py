import os
import glob
import time
import outlook_handler


class DownloadFolderWatcher:
    def __init__(self, callback_on_file_found):
        self.watch_dirs = [
            os.path.join(os.path.expanduser('~'), 'Downloads'),
            os.path.join(os.path.expanduser('~'), 'Desktop'),
        ]
        # 아웃룩 임시 폴더 추가
        try:
            outlook_dirs = outlook_handler.get_outlook_temp_dirs()
            for od in outlook_dirs:
                if od and od not in self.watch_dirs:
                    self.watch_dirs.append(od)
        except Exception:
            pass

        self.callback = callback_on_file_found
        self.known_files = set(self._get_current_files())

    def _get_current_files(self):
        files = []
        for d in self.watch_dirs:
            if not os.path.exists(d):
                continue
            for ext in ('*.pdf', '*.html', '*.htm'):
                for fp in glob.glob(os.path.join(d, ext)):
                    files.append(fp)
        return files

    def check_new_files(self):
        current_files = set(self._get_current_files())
        new_files = current_files - self.known_files
        if new_files:
            self.known_files = current_files
            for file_path in new_files:
                if self.callback:
                    self.callback(file_path)

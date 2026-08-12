import os
import time

class DownloadFolderWatcher:
    def __init__(self, callback_on_pdf_found):
        self.download_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        self.callback = callback_on_pdf_found
        self.known_files = set(self._get_current_pdfs())

    def _get_current_pdfs(self):
        if not os.path.exists(self.download_dir):
            return []
        files = []
        for f in os.listdir(self.download_dir):
            if f.lower().endswith('.pdf'):
                files.append(os.path.join(self.download_dir, f))
        return files

    def check_new_files(self):
        current_pdfs = set(self._get_current_pdfs())
        new_pdfs = current_pdfs - self.known_files
        if new_pdfs:
            self.known_files = current_pdfs
            for pdf_path in new_pdfs:
                if self.callback:
                    self.callback(pdf_path)

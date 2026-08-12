import os
import tempfile
import sys

def fetch_outlook_attachments():
    if sys.platform != 'win32':
        return []

    try:
        import win32com.client
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32com.client.Dispatch("Outlook.Application")

        selected_items = []

        # 1. 아웃룩 탐색기(Explorer)에서 선택된 이메일
        try:
            explorer = outlook.ActiveExplorer()
            if explorer and explorer.Selection.Count > 0:
                for i in range(1, explorer.Selection.Count + 1):
                    selected_items.append(explorer.Selection.Item(i))
        except Exception as e:
            print(f"Selection fetch err: {e}")

        # 2. 열려 있는 이메일 창(Inspector)
        if not selected_items:
            try:
                inspector = outlook.ActiveInspector()
                if inspector and inspector.CurrentItem:
                    selected_items.append(inspector.CurrentItem)
            except Exception as e:
                print(f"Inspector fetch err: {e}")

        # 3. 선택/열린 메일이 없으면 받은편지함(Inbox) 최신 10개 메일 스캔
        if not selected_items:
            try:
                namespace = outlook.GetNamespace("MAPI")
                inbox = namespace.GetDefaultFolder(6) # 6 = olFolderInbox
                messages = inbox.Items
                messages.Sort("[ReceivedTime]", True)
                for i in range(1, min(11, messages.Count + 1)):
                    selected_items.append(messages.Item(i))
            except Exception as e:
                print(f"Inbox scan err: {e}")

        extracted_files = []
        temp_dir = tempfile.gettempdir()

        for item in selected_items:
            try:
                if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                    for att in item.Attachments:
                        fname = att.FileName
                        if fname.lower().endswith('.pdf'):
                            save_path = os.path.join(temp_dir, f"outlook_{fname}")
                            att.SaveAsFile(save_path)
                            extracted_files.append(save_path)
            except Exception as e:
                print(f"Item attachment err: {e}")

        return extracted_files
    except Exception as e:
        print(f"Outlook fetch error: {e}")
        return []

if __name__ == '__main__':
    res = fetch_outlook_attachments()
    print(f"Extracted {len(res)} files: {res}")

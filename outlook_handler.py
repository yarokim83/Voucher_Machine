import os
import tempfile
import sys

def fetch_outlook_attachments():
    if sys.platform != 'win32':
        return []

    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        explorer = outlook.ActiveExplorer()
        
        selected_items = []
        if explorer and explorer.Selection.Count > 0:
            for i in range(1, explorer.Selection.Count + 1):
                selected_items.append(explorer.Selection.Item(i))
        else:
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)
            for i in range(1, min(6, messages.Count + 1)):
                selected_items.append(messages.Item(i))

        extracted_files = []
        temp_dir = tempfile.gettempdir()

        for item in selected_items:
            if hasattr(item, 'Attachments') and item.Attachments.Count > 0:
                for att in item.Attachments:
                    fname = att.FileName
                    if fname.lower().endswith('.pdf'):
                        save_path = os.path.join(temp_dir, f"outlook_{fname}")
                        att.SaveAsFile(save_path)
                        extracted_files.append(save_path)

        return extracted_files
    except Exception as e:
        print(f"Outlook fetch error: {e}")
        return []

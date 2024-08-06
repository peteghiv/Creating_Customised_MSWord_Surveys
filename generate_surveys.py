import os, csv, win32com.client

SURVEY_PASSWORD = 'myP@ssw0rd1234'
cwd = os.getcwd()
TEMPLATE_PATH = os.path.join(cwd, 'Survey_Template.docx')

def add_basic_info(word: win32com.client, doc_path: str, basic_info: dict):
    # Open template document
    doc = word.Documents.Open(TEMPLATE_PATH)
    doc.SaveAs(doc_path)

    # Fill in basic information
    for key, value in basic_info.items():
        # Assuming you use bookmarks for filling information
        try:
            # Get the bookmark
            bookmark = doc.Bookmarks(key)
            bookmark_range = bookmark.Range

            # Temporarily remove the bookmark
            bookmark.Delete()

            # Update the range text
            bookmark_range.Text = value

            # Re-add the bookmark at the same range
            doc.Bookmarks.Add(key, bookmark_range)
        except:
            print(f"Bookmark '{key}' not found in the document.")
    
    # Protect the document to for form filling
    doc.Protect(Type=2, NoReset=False, Password=SURVEY_PASSWORD)

    # Save the document
    doc.Save()
    doc.Close()

def read_csv(csv_path: str):
    # Open CSV file
    with open(csv_path) as f:
        reader = csv.reader(f)
        header = next(reader)

        # Consolidate all rows into a dict
        result = []
        for row in reader:
            temp = {}
            for i in range(len(header)):
                temp[header[i]] = row[i]
            result.append(temp)

    return result

def main():
    # Get the details for the surveys to be created
    data = read_csv(os.path.join(cwd, 'survey_takers.csv'))

    # Start MS Word
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    for item in data:
        filename = f'{item["company"]}_{item["name"]}.docx'
        file_path = os.path.join(cwd, 'Generated_Surveys', filename)
        add_basic_info(word, file_path, item)

    # Close MS Word
    word.Quit()

if __name__ == '__main__':
    main()
    print('Done')
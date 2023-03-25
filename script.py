import win32com.client
from transformers import pipeline
import pandas as pd

# set up zero-shot classification pipeline
classifier = pipeline("zero-shot-classification")

# set up Outlook client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

df = pd.DataFrame()
def iterate(df):
    # retrieve inbox and subfolders
    inbox = outlook.GetDefaultFolder(6)
    subfolder_names = [folder.Name for folder in inbox.Folders]
    subfolders = {name: inbox.Folders[name] for name in subfolder_names}
    emails = inbox.Items

    for email in emails:
        # classify email using zero-shot classification model
        classification = classifier(email.Subject + " " + email.Body[:255], list(subfolders.keys()))
        label = classification["labels"][0]

        # move email to appropriate subfolder
        destination_folder = subfolders[label]
        email.Move(destination_folder)
        df = pd.concat([df, pd.DataFrame(classification).head(1)])

        email_count = len(inbox.Items)
        if email_count > 0:
            iterate(df)
        else:
            return df
iterate(df)
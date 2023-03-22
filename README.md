# Outlook-Email-Sorter
This is an add-in that uses a language model for text classification of the emails in your inbox using the Inbox's subfolder names as the class labels.

Main features/goals of the intended add in  are as follows:

1. activation of script on email inbox using a button to sort emails in inbox (batch)
2. option to activate script upon receipt of email (stream)
4. option to have language model downloaded
5. class labels derived from Inbox subfolder names
6. zero-shot classification of text body and subject
7. change location of email to the relevant inbox subfolder if confidence above set threshold (option?)
8. maintain unread status of email 


options

ONNX format of model used to avoid python

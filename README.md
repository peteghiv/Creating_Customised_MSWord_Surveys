# Creating Customised Microsoft Word Surveys

This builds on my previous repository ([Automating MS Word Surveys](https://github.com/peteghiv/Automating_MS_Word_Surveys)) by customising the MS Word survey to each user.

## Notable changes from the previous repository:
1. **Use of Content Controls rather than Legacy Form Fields.** This allows compulsory fields to be highlighted, but subsequently display normal text once an input is in. Hence, it reduces the likelihood of an erroneous survey submission.

2. **Use of Bookmarks to store survey respondent information.** This is useful in the event that you have a database of survey respondents and want to ensure the accuracy of inputs, preventing typographical errors.

## Brief Explanation of Code
For each survey, the code will create a new Word document, update the respondent information, and then lock the document for form-filling only.

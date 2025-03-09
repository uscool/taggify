# taggify
"Taggify," is a desktop-based email organization tool that automates email categorization using keywords and assigns priority levels based on content urgency. This tool, implemented in Python, is designed for users who need a streamlined way to organize, access, and manage their emails by tagging and prioritizing based on the email content.
## Functionality
1.	Fetches emails from a specified Outlook inbox.
2.	Categorizes emails by matching keywords in the subject or body to predefined tags, such as "work," "important," or "urgent."
3.	Displays emails in a sortable and filterable table for easy access. 
4.	Opens emails directly in Outlook from the application. 
5.	Allows detagging of emails that are no longer relevant, storing this information to prevent future re-tagging. 
## How It Works 
The program uses the win32com library to access the Outlook inbox and fetch emails, extracting their subject, received time, and body content. Using regular expressions from the re module, Taggify matches specific keywords in each email to predefined tags stored in the tags dictionary. Each tag reflects categories like urgency, importance, or the type of content (e.g., work-related or class-related emails). 
The tool then: 
1.	Prioritizes emails based on the tag categories, with "urgent" emails receiving the highest priority, followed by "important" and then "normal." 
2.	Displays emails in a tkinter GUI table (using ttk.Treeview), where each email’s subject, received time, and tags are shown. Users can sort and filter emails using interactive buttons. 

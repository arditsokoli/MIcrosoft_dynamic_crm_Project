# Assignment for Dynamics CRM:

### Business Case: 
We need a solution to solve our daily task to import invoices on system, they are in Excel format and contain all information needed.
### 📕 Information :
#### fisoft_product
- Name
- Price
- Product code
#### fisoft_invoice :
- Description
- Number
- Created Date
- Customer (lookup)
- Total Amount
- Paid
- Paid Date
#### fisot_invoice_product :
- Product(lookup)
- Invoice(lookup)
- Unit
- Quantity
- Amount
#### fisoft_customer :
- Name
- Address
- Revenue
- Email 

### This is the flow: 

![Blockscema](../../tree/master/img/bllokskema.png)

1. Create a console application to import invoices to CRM 365
- File is type Excel and stored on c:\FISOFT_Consulting\invoices\  on your local machine
- If an error is thrown during execution, it should be saved in a text file on your local machine under c:\FISOFT_Consulting\Error_Logs\  
2. When you open a customer record you should see Name, Address, Revenue and all his invoices (Revenue is total amount from all his invoices)
- If Revenue is greater than 1000Euro show am info message on form: “Top Customer”
3. When you open Invoice it should show Description,Created Date, Customer, Total Amount and all invoice products on grid 
- By default paid field is No, and user can change it manually.
- If paid is No show an warning: ”This invoice {Invoice Number} need Attention”
- When user changes to yes this field became “Read Only” with  Paid Date
- Paid Date is calculated automatically with current date.
4. Product and Invoice Product should have on form only custom fields on document
5. When a new Customer is Created an email should be sent to him to congrats him for being part of Company.
- Revenue field should be visible only for Customer that address contains Albania

#### Programing Language:
Javascript
C#

Contact me:  [ardit.sokoli@fshnstudent.info](mailto:ardit.sokoli@ap.edu.al?subject=[GitHub]%20Source%20Han%20Sans)

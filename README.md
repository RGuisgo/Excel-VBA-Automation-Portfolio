For the BP EUROPA file, it automates clearing of the payable netted  BP invoice that are due for payment afer posting these invoices.
For the BP EUROPA SE file, it automates making of BP invoice that are due for payment afer posting these invoices in to one file with both the sales and purchases side and these side are netted into whether there is a payable netting, zero netting or receivable netting.
For the BP EUROPA (1) file, it automates clearing of the receivale netted  BP invoice that are due for receipt afer posting these invoices.
For the BP EUROPA (2) file, it automates clearing of the zero netted  BP invoice that are cancelling each other because there same amount on the sales side and purchase side afer posting these invoices.

For the IC-DA macro, it automates the formatting of export from SAP into a well defined excel workbook without any unnecessary columns and rows.

For the Truck robot, it automates the posting of both invoices with or without purchasing order numbers that are due directly for payment as soon as a purchase is done.

For the buy sell contract macro, it automates the extracting of both sales and purchase contract numbers from a large contract admin file and it helps to help duplicate contract numbers from the file.

For the price change macro, it formats raw data from SAP, go through two extra excel workbooks to look for matching trade numbers and insert vendor names based on the trade numbers found.

All these files script have SAP GUI scripting in them since invoices posting and netting clearing are both done in SAP system. The importance of these macros are to save time, and make invoice processing and clearing easy and efficient.

## How to Use
1. Except for the buy sell contract macro, all the others are connected with SAP. So, open SAP and go into the required system you will need your invoices posted or cleared.
2. There will be no need to insert a transaction code in SAP, as they are already embeded in the macro.
3. Open the file in Excel and enable macros.
4. Run the macro:
   -  click **Run**.
   -  
NB* For cases where you need to use this macro and the transaction codes in SAP are not the same, go into the macro and do the changes with the correct transaction code you need.
  

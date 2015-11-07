The following form is a simple Excel macro enabled sheet
that allows you to do several integrity checks, provides a workflow
and allows tracking.  

There are 3 subroutines that correspond to each buttom. 

In addition, the code in the main page allows the user
to password protect hidden worksheets. 


A summary of the subroutines


Sub enter()

1) Checks for missing values on key fields
2) Checks that there are at least 2 entries and no more than 200 entries in the form
3) Message box to confirm choice 
4) Some fields are always negative. It checks to see if field was entered as positive
   and automatically transforms it to negative 
3) Calls VBA_to_append_existing_text_file subroutine


Sub VBA_to_append_existing_text_file subroutine()
1) Checks to see if item is in sample (From tab Balances)
2) Select file and append it to Balances.txt on ROOT\OUTPUT 
3) Create an individual file and place it in ROOT\OUTPUT\Indidividual
4) Mark Balance as completed 

sub Clear()
1) Clear all fields in range


Sub  getnextbalance()
1) Clear contents 
2) Search for next balance in line (marked by a 0 in BALANCES tab)
3) Populate balance and year to sheet


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
1) Password protect hidden tabs 
# VBA-challenge
By Elee Saleem

Multiple_year_stock_data.xlsm unfortunately can't be uploaded since github doesn't allow files bigger than 25mb to be uploaded on its platform and this multiple year file I tried to upload is 95.6 mb

References:
*Most of codes were inspired from class 3 exercise 6 and 7 and also from class 2 exercise 6

*Except :
ws.Range("k" & table2row).NumberFormat = "0.00%" this line was inspired from:
statology.org/vba-percentage-format/

Application.Max(ws.Range("K2:K" & Lastrow3).Value)
Application.Min(ws.Range("K2:K" & Lastrow3).Value) these 2 lines were inspired from :
mrexcel.com and stackoverflow.com I had to read couple of pages both of them

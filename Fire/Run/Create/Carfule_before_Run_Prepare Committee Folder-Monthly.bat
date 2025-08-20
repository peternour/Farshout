@echo off
set CUR_DATE=%DATE:~10,4%-%DATE:~4,2%-%DATE:~7,2%
set CUR_Mo=%DATE:~10,4%-%DATE:~4,2%
mkdir "D:\Committee\01 Inform"
mkdir "D:\Committee\01 Inform\Reports"
mkdir "D:\Committee\02 Com-Reports"
mkdir "D:\Committee\02 Com-Reports\Card_Holder_Sheet"
mkdir "D:\Committee\03 Disburise"
mkdir "D:\Committee\04 Reinforcement"
mkdir "D:\Committee\05 Neg-List"
mkdir "D:\Committee\05 Neg-List\05-01 Clients"
mkdir "D:\Committee\05 Neg-List\05-02 Guarantors"
mkdir "D:\Committee\06 Helping-Files"
mkdir "D:\Committee\07 Clients-Contract"
mkdir "D:\Committee\08 Sub-Com"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Ahly.xlsm" "D:\Committee\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%_Ahly.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Edy.xlsm" "D:\Committee\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%_Edy.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Misr_Mid.xlsm" "D:\Committee\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%_Misr_Mid.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_Misr_short.xlsm" "D:\Committee\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%_Misr_short.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Meza_Trans_spanish.xlsm" "D:\Committee\02 Com-Reports\Card_Holder_Sheet\%CUR_DATE%_spanish.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ta3zez.docx" "D:\Committee\04 Reinforcement\%CUR_DATE%_Ta3zez.docx"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Ma7dar.xlsm" "D:\Committee\02 Com-Reports\%CUR_DATE%_Ma7dar.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Iscour_Report.xlsm" "D:\Committee\01 Inform\Reports\%CUR_DATE%_Iscour_Report.xlsm"
copy "D:\COMMITTEE FORMES\Miza-Monthly\Co_py\Cards_Rearrange.xlsm" "D:\Committee\02 Com-Reports\%CUR_DATE%_Cards_Rearrange.xlsm"
copy "D:\COMMITTEE FORMES\Neg_List_Form\Neg_List_Report.xlsx" "D:\Committee\05 Neg-List\%CUR_DATE%_Neg_List_Report.xlsx"
copy "D:\COMMITTEE FORMES\Neg_List_Form\Neg_List_Report.docx" "D:\Committee\05 Neg-List\%CUR_DATE%_Neg_List_Report.docx"
copy "D:\COMMITTEE FORMES\Miza_Cards\Miza_Daiy_Disburse\Miza_Recevied_Report\Misr\Miza_Received_Report.docx" "D:\Committee\02 Com-Reports\%CUR_DATE%_Miza_Received_Report.docx"










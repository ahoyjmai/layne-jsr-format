
#[Column name, number formatting, (+)value, (-)value, 'prev' if (-) is to be subtracted from previous month]
# column letters correspond to JSR downloaded from BI
#       ["Column title seen in JSR"                     ,"Formatting"      ,"Value to grab from Original All", "Value to subtract from Original All", "Type in Prev if you want the subtract value to be from Original Prev instead of All"],
HEADERMAP=[
        ["Contract Type"                                ,"General"         ,"B","",     ""],
        ["S"                                            ,"General"         ,"","",      ""] ,
        ["Job #"                                        ,"General"         ,"D","",     ""],
        ["Parent #"                                     ,"General"         ,"E","",     ""],
        ["Job Name"                                     ,"General"         ,"F","",     ""],
        ["Est Sales (Contract)"                         ,"#,##0_);(#,##0)" ,"K","",     ""],
        ["E1 Forecasted Sales (New)"                    ,"#,##0_);(#,##0)" ,"Q","",     ""],
        ["Sales Delta Forecast vs Contract (New)"       ,"#,##0_);(#,##0)" ,"Q","K",    ""],
        ["Total Revenue (Hypo.)"                        ,"#,##0_);(#,##0)" ,"L","",     ""],
        ["Monthly Revenue (Hypo.)"                      ,"#,##0_);(#,##0)" ,"L","L",    "prev"],
        ["Monthly Billings"                             ,"#,##0_);(#,##0)" ,"N","N",    "prev"],
        ["Est Cost (Contract)"                          ,"#,##0_);(#,##0)" ,"T","",     ""],
        ["E1 Forecasted Cost (New)"                     ,"#,##0_);(#,##0)" ,"AN","",    ""],
        ["Cost Delta Forecast vs Contract (New)"        ,"#,##0_);(#,##0)" ,"AN","T",   ""],
        ["Actual Total Cost"                            ,"#,##0_);(#,##0)" ,"AB","",    ""],
        ["Actual Monthly Cost"                          ,"#,##0_);(#,##0)" ,"U","",     ""],
        ["Act Total Cost incl 995 & T&D"                ,"#,##0_);(#,##0)" ,"","",      ""],  #
        ["Act Monthly Cost incl 995 & T&D"              ,"#,##0_);(#,##0)" ,"","",      ""],    # Orig All "U" + However JSR col S is calculated
        ["Actual Total Margin incl 995"                 ,"0.0%_);(0.0%)" ,"","",      ""],  #
        ["YTD Hourly Manhours"                          ,"#,##0_);(#,##0)" ,"","",      ""],    # would like date in col title  # Map to new file tab1 col J
        ["Est Accruals for 995 and T&D"                 ,"#,##0_);(#,##0)" ,"","",      ""],    # JSR Col R * New File tab2 Col G
        ["Est Margin (Forecasted)"                      ,"#,##0_);(#,##0)" ,"AX","",    ""],
        ["Est Margin % (Forecasted)"                    ,"0.0%_);(0.0%)"   ,"AY","",    ""],
        ["Actual Margin"                                ,"#,##0_);(#,##0)" ,"AV","",    ""],
        ["Actual Margin %"                              ,"0.0%_);(0.0%)"   ,"AW","",    ""],
        ["POC %"                                        ,"0.0%_);(0.0%)"   ,"H","",     ""],
        ["Billings"                                     ,"#,##0_);(#,##0)" ,"M","",     ""],
        ["POC Receivable"                               ,"#,##0_);(#,##0)" ,"L","M",    ""],
        ["Trade AR"                                     ,"#,##0_);(#,##0)" ,"O","",     ""],
        ["Open Retainage"                               ,"#,##0_);(#,##0)" ,"P","",     ""],
        ["Billings > Sales"                             ,"#,##0_);(#,##0)" ,"M","L",    ""],
        ["Sales > Billings"                             ,"#,##0_);(#,##0)" ,"L","M",    ""],
        ["Billings > Cost"                              ,"#,##0_);(#,##0)" ,"AI","",    ""],
        ["Cost > Billings"                              ,"#,##0_);(#,##0)" ,"AJ","",    ""],
        ["Orig Cntrct Amt"                              ,"#,##0_);(#,##0)" ,"I","",     ""],
        ["Orig Cntrct Margin"                           ,"#,##0_);(#,##0)" ,"AO","",    ""],
        ["Orig Cntrct Margin %"                         ,"0.0%_);(0.0%)"   ,"AP","",    ""],
        ["Monthly Est Sales Change"                     ,"#,##0_);(#,##0)" ,"K","K","prev"],
        ["Monthly Est Cost Change"                      ,"#,##0_);(#,##0)" ,"T","T","prev"],
        ["Monthly Est Profit Change"                    ,"#,##0_);(#,##0)" ,"AU","",    ""],
        ["Actual Margin Monthly Change"                 ,"#,##0_);(#,##0)" ,"AV","AV","prev"],
        ["Forecast Margin Monthly Change"               ,"#,##0_);(#,##0)" ,"AX","AX","prev"],
        ["LEMSO ITD Labor Actual"                       ,"#,##0_);(#,##0)" ,"AC","",    ""],
        ["LEMSO ITD Equipment Actual"                   ,"#,##0_);(#,##0)" ,"AD","",    ""],
        ["LEMSO ITD Materials Actual"                   ,"#,##0_);(#,##0)" ,"AF","",    ""],
        ["LEMSO ITD Subcontracts Actual"                ,"#,##0_);(#,##0)" ,"AG","",    ""],
        ["LEMSO ITD Hauling Costs Actual"               ,"#,##0_);(#,##0)" ,"AE","",    ""],
        ["PM"                                           ,"General"         ,"G","",     ""],
        ["SP"                                           ,"General"         ,"","",      ""],
        ["Area"                                         ,"General"         ,"BC","",    ""],
        ["Cost Cntr Finance (RBU)"                      ,"General"         ,"","",      ""],
        ["Cost Cntr Home"                               ,"General"         ,"BE","",    ""],
        ["Plan Start Date"                              ,"m/d/yyyy"        ,"BI","",    ""],
        ["Actual Start Date"                            ,"m/d/yyyy"        ,"BK","",    ""],
        ["End Date"                                     ,"m/d/yyyy"        ,"BL","",    ""],
        ["Product Line"                                 ,"General"         ,"BH","",    ""],
        ["Customer"                                     ,"General"         ,"","",      ""],
        ["Bid Type"                                     ,"General"         ,"BO","",    ""],
        ["Prime/Sub"                                    ,"General"         ,"BM","",    ""],
]

# LMS_ETL
Consume and compile reports from Training LMS

Dependencies:
  openpyxl
  pypiwin32
  

Edit Lines 44 & 45 to specify user name in path1 and path2 var.


With IC region_list:

                  CCC      CNR     IC      NER   NWR    SCR    SER    SNR
                    1       2       3       4     5       6     7       8

W/O IC region_list:

                      CCC   CNR    NER    NWR    SCR    SER    SNR
                      1       2     3       4     5       6      7

Steps:

1: Select Start Menu and type 'cmd'
2: Open Command Prompt
3: enter 'cd C:\Users\<your AD Account name>\LMS_ETL_temp'
4: Press enter
5: enter 'python LMS_ETL.py'
6: Press enter
7: Profit!!

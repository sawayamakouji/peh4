The article "VBAを使って商品の売上を自動で取得させるとガチ便利だった件" on Qiita describes how to automate the process of acquiring product sales data using VBA and ChatGPT. The author, a beginner in VBA, used ChatGPT to gener[1]ate a macro that transfers sales quantity and amount from an "単品実績" (individual item performance) sheet to a "商品名" (product name) sheet in Excel, rounding the sales amount to the nearest thousand yen.

The process involves:
1.  Creating an Excel format with three she[1]ets: "計画書" (plan sheet), "商品名" (product name), and "単品実績" (individual item performance).
2.  Using JAN codes to match products between the "商品名" and "単品実績" sheets.
3.  Having [1]ChatGPT generate the VBA code by uploading the Excel file and asking a general question.
[1]4.  Pasting the generated VBA code into a standard module in the Excel VBA editor.
5.  Adding a butt[1]on to the "商品名" sheet to execute the macro.

The author received positive feedback [1]from a colleague regarding the time-saving aspect of the macro, but[1] also noted that manually entering JAN codes remains a bottleneck. The article concludes by highlighting VBA as an accessible RPA (Robotic Process Automation) tool, especial[1]ly when combined with generative AI like ChatGPT, for internal business process improvement.[1]

Sources:
[1] VBAを使って商品の売上を自動で取得させるとガチ便利だった件 - Qiita (https://qiita.com/horigome2245/items/86781d7f933c4ae644f6)
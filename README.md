# PDF-Molecular-Pathology-Laboratory-Report-recognition
金域医学检验实验室实验报告PDF识别
表格具有同一性，同时也具有多样性
识别主要分为两方面
一 ，识别用户信息
二， 识别检测结果

就用户信息来看，有的用户信息过少，不足以填充满整个表格，导致pdfplumber无法识别其为表格，
不过幸好前面有写正则匹配，直接在text中查找，来填补信息的不足。

就检测结果来看，检测结果分为三部分，都能顺利识别，这很好。但是分布的并不固定，好在都在前两页，
就前两页进行检查即可。

幸好前面程序有Parsing_userinfo(page01,each_pdfinfo)，Parsing_type(page01) 两个正则函数，
极大方便了检测和识别。

不过程序待完善的地方还很多，有时间重构一下

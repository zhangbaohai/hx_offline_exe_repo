v2.3.5c 修复：把 PayrollTab/TransferTab/LibraryTab 全部并入 app_exact.py，避免 NameError。
运行：pip install -r requirements.txt && python app_exact.py
功能：行号库导入(txt/xls/xlsx/csv)；新增编辑；批量导入校验；自动识别“华夏银行”设置转账方式；行别信息类型可为空。

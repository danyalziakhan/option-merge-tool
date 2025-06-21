chcp 65001
@echo off
rem get current this batch file directory
set dir=%~dp0

CALL .venv\Scripts\activate
CALL .venv\Scripts\python run.py --input_file INPUT_FILE.xlsx --template_file TEMPLATE_FILE_DB.xlsx --first_column "원본 상품명" --second_column "원가\n[필수]" --join_by "," --output_column "옵션상세명칭(1)\n[사방넷]" --column_to_dropna "원본 상품명" --columns_to_drop_dulicates "원본 상품명,옵션상세명칭(1),물류처ID,모델NO"

pause
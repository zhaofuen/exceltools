import os
from datetime import datetime
import streamlit as st
# import xlwings as xw
# import excel3img
from openpyxl import load_workbook, Workbook
from PIL import ImageDraw

def split_data(src, target_folder):
    # # 检查源文件路径是否存在
    # if not os.path.exists(src):
    #     print('文件路径不正确，请检查')
    #     return
    
    # # 确保目标文件夹存在
    # os.makedirs(target_folder, exist_ok=True)
    
    try:
        # 加载工作簿
        workbook = load_workbook(filename=src)
        
        for sheet in workbook:
            # 处理工作表名称，避免非法字符
            safe_name = ''.join([c for c in sheet.title if c.isalpha() or c.isdigit() or c == ' ']).rstrip()
            
            # 创建一个新的工作簿
            new_workbook = Workbook()
            new_sheet = new_workbook.active
            new_sheet.title = sheet.title
            
            # 复制工作表内容
            for row in sheet.iter_rows(values_only=True):
                new_sheet.append(row)
            
            # 保存新的工作簿
            new_workbook_path = os.path.join(target_folder, f"{safe_name}.xlsx")
            new_workbook.save(new_workbook_path)
            print(f"工作表 {sheet.title} 已保存到 {new_workbook_path}")
    
    except Exception as e:
        print(f"错误信息: {e}")

# 示例调用
# split_data("example.xlsx", "./报表/")
def delete_empty_folders_in_current_directory():
        """
        检查当前目录下是否有空的文件夹，并删除这些空文件夹。
        """
        current_directory = os.getcwd()  # 获取当前目录路径
        
        for root, dirs, files in os.walk(current_directory, topdown=False):
            for dir_name in dirs:
                folder_path = os.path.join(root, dir_name)
                if not os.listdir(folder_path):  # 检查文件夹是否为空
                    try:
                        os.rmdir(folder_path)
                        print(f"已删除空文件夹: {folder_path}")
                    except Exception as e:
                        print(f"删除文件夹 {folder_path} 时出错: {e}")
def create_directory(tool): 
        delete_empty_folders_in_current_directory()
        # 获取当前时间，并格式化为字符串  
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  
    
        # 指定要创建的目录的路径（例如，在当前工作目录下）  
        directory_path = os.path.join(os.getcwd(), current_time+tool)  
        
        # 创建目录  
        try:  
            os.mkdir(directory_path)  
            print(f"目录 '{directory_path}' 创建成功")  
        except OSError as error:  
            print(f"创建目录 '{directory_path}' 失败。错误：{error}")
        return directory_path

def download_res(file_path):
        if file_path:
            #下载
            with open(file_path,'rb') as file:
                btn = st.download_button(
                    label="📥下载文件",
                    data=file,
                    file_name= file_path.split('/')[-1],
                    mime='text/xlsx'
                )
        else:
            print('文件不存在')
def download_zip_file(zip_file_path):
    """
    提供下载 ZIP 文件的功能。

    :param zip_file_path: ZIP 文件的路径
    """
    # 读取 ZIP 文件内容
    with open(zip_file_path, "rb") as f:
        zip_data = f.read()

    # 创建下载按钮
    st.download_button(
        label="下载 ZIP 文件",
        data=zip_data,
        file_name=os.path.basename(zip_file_path),
        mime="application/zip"
    )


languages = {
    "CN":{
        "button":"浏览文件",
        "instructions":"将文件拖放到此处",
        "limits":"每个文件限制200MB",
    },
    # "EN":{
    #     "button":"Browse Files",
    #     "instructions":"Drag and drop files here",
    #     "limits":"Each file limited to 200MB",
    # },
}

lang = None

def style_language_uploader():
    language = 'CN'
    hide_label = (
        """
    <style>
        div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"] {
           color: white;
        }
        div[data-testid="stFileUploader"]>section[data-testid="stFileUploaderDropzone"]>button[data-testid="baseButton-secondary"]::after {
            content: "BUTTON_TEXT";
            color:black;
            display: block;
            position: absolute;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>span {
           visibility:hidden;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>span::after {
           content:"INSTRUCTIONS_TEXT";
           visibility:visible;
           display:block;
        }
         div[data-testid="stFileUploaderDropzoneInstructions"]>div>small {
           visibility:hidden;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"]>div>small::before {
           content:"FILE_LIMITS";
           visibility:visible;
           display:block;
        }
    </style>
     """.replace("BUTTON_TEXT", languages.get(language).get("button"))
        .replace("INSTRUCTIONS_TEXT", languages.get(language).get("instructions"))
        .replace("FILE_LIMITS", languages.get(language).get("limits"))
    )

    st.markdown(hide_label, unsafe_allow_html=True)


# def split_data(src,target_folder):    
#     # if not os.path.exists(src):        
#     #     print('文件路径不正确，请检查')        
#     #     return    
#     # target_folder = './报表/'    
#     # os.makedirs(target_folder, exist_ok=True)    
#     # 启动Excel应用，不显示界面    
#     app = xw.App(visible=False, add_book=False)    
#     try:        
#             # 加载工作簿        
#             workbook = app.books.open(src)        
#             for sheet in workbook.sheets:            # 处理工作表名称，避免非法字符            
#                  safe_name = ''.join([c for c in sheet.name if c.isalpha() or c.isdigit() or c == ' ']).rstrip()
#                  workbook_split = app.books.add()            
#                  sheet_split = workbook_split.sheets[0]            
#                  sheet.api.Copy(Before=sheet_split.api)            
#                  workbook_split.save(os.path.join(                
#                  target_folder, f"{safe_name}.xlsx"))            
#                  workbook_split.close()    
#     except Exception as e:        
#         print(f"错误信息: {e}")    
#     finally:        # 关闭工作簿        
#         workbook.close()        # 关闭Excel实例        
#         app.quit()

# def batch_process_excel_files(folder_path):    
#      print("批量处理文件夹下所有Excel文件...")    # 遍历文件夹中的所有文件    
#      for file in os.listdir(folder_path):        
#         if file.endswith(('.xls', '.xlsx', '.xlsm')):           
#              file_path = os.path.join(folder_path, file)            
#              split_data(file_path)

# file_path = './data/'
# batch_process_excel_files(file_path)
# Excel 转换为图片的函数
# def out_img(excel_file, sheet_list,outputdir):
#     """
#     将Excel文件中的指定工作表转换为图片。

#     参数:
#     excel_file: string, Excel文件的路径。
#     sheet_list: list, 需要转换为图片的工作表名称列表。
#     outputdir: string, 输出图片文件的目录路径。
#     截图效果非常好。
#     返回:
#     无返回值，但会在当前目录下生成对应工作表的图片文件。
#     """
#     try:
#         # 开始转换操作前的提示
#         print("开始截图，请耐心等待....")
#         # 遍历工作表列表，对每个工作表进行转换
#         for sheet_name in sheet_list:
#             # 调用excel2img模块的export_img函数进行转换
#             # 参数包括Excel文件路径、图片文件名、工作表名和自定义配置（这里设为None）

#             image_filname = os.path.join(outputdir,f"{sheet_name}.png")
#             excel3img.export_img(excel_file, image_filname,sheet_name, None)
#             print(f"{sheet_name} 截图完成")
#     except Exception as e:
#         # 捕获转换过程中可能出现的异常，并打印异常信息
#         print("截图失败", e)
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO

def out_img(excel_file, sheet_list, outputdir):
    """
    将Excel文件中的指定工作表转换为图片。

    参数:
    excel_file: string, Excel文件的路径。
    sheet_list: list, 需要转换为图片的工作表名称列表。
    outputdir: string, 输出图片文件的目录路径。
    返回:
    无返回值，但会在指定目录下生成对应工作表的图片文件。
    此函数只能搞定了内容转存并截图，丧失了原来的各种格式。
    """
    try:
        # 开始转换操作前的提示
        print("开始截图，请耐心等待....")
        
        # 确保输出目录存在
        if not os.path.exists(outputdir):
            os.makedirs(outputdir)
        
        # 加载工作簿
        workbook = load_workbook(filename=excel_file)
        
        # 遍历工作表列表，对每个工作表进行转换
        for sheet_name in sheet_list:
            # 获取指定的工作表
            sheet = workbook[sheet_name]
            
            # 创建一个空白的图片
            img = PILImage.new('RGB', (1000, 1000), color = (255, 255, 255))
            draw = ImageDraw.Draw(img)
            
            # 获取工作表的行和列
            for row in sheet.iter_rows():
                for cell in row:
                    # 获取单元格的值
                    cell_value = cell.value
                    if cell_value is not None:
                        # 在图片上绘制单元格内容
                        draw.text((cell.column * 100, cell.row * 50), str(cell_value), fill="black")
            
            # 保存图片
            image_filename = os.path.join(outputdir, f"{sheet_name}.png")
            img.save(image_filename)
            
            print(f"{sheet_name} 截图完成")
    
    except Exception as e:
        # 捕获转换过程中可能出现的异常，并打印异常信息
        print("截图失败", e)

# 示例调用
# out_img("example.xlsx", ["Sheet1", "Sheet2"], "./output_images")
# 导入所需的库
import streamlit as st
import openpyxl


# 创建一个Streamlit页面
def main():
    # 设置页面标题
    work_hours = {name: 0 for name in [
            "贾学涛", "王亚伦", "李春荣", "薛鹏", "兰梅芳", "管燕坤", "赖鑫源", "方昊", "罗华松",
            "庄建柳", "李彤", "朱晓阳", "徐苏哲", "张建全", "张慧慧", "柯文欣", "陈紫瑶", "贾玲玲",
            "杜郝", "陈昭玉", "袁博艺", "杨金金", "伍龙秀", "许达", "于永康", "郑佳豪"
        ]}
    st.title('简单的值班表工时统计')

    # 创建一个文件上传器
    uploaded_file = st.file_uploader("选择要统计的值班表", type=['xlsx', 'xls'])

    # 检查是否有文件被上传
    if uploaded_file is not None:
        # 读取Excel文件
        workbook = openpyxl.load_workbook(uploaded_file)

        # 选择第一个工作表
        sheet = workbook.worksheets[0]




        # 定义遍历的范围，这里以A2:N7为例，你可以根据需要修改
        start_row, end_row = 4, 45
        start_col, end_col = 3, 13  # A是1，B是2，以此类推

        # 定义权重字典，根据列号来分配权重
        weight_dict = {
            3: 2,  # 3列（C列）权重为2
            5: 2,  # 5列（E列）权重为2
            7: 1,  # 7列（H列）权重为1
            11: 1,  # 11列（L列）权重为1
            9: 1.5,  # 9列（I列）权重为1.5
            13: 1.5  # 13列（N列）权重为1.5
        }

        # 遍历指定范围内的单元格
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                # 检查单元格是否包含姓名
                if cell.value in work_hours:
                    # 获取权重，如果列不在权重字典中，默认权重为1
                    weight = weight_dict.get(col, 1)
                    # 累加工时到字典
                    work_hours[cell.value] += weight
    col1, col2 = st.columns(2)
    with col1:
        st.subheader('姓名')
        for name in work_hours.keys():
            st.write(name)
    with col2:
        st.subheader('工时')
        for value in work_hours.values():
            st.write(value)

# 运行主函数
if __name__ == "__main__":
    main()

import pandas as pd


def generate_html_from_excel(excel_file, html_file):
    # 读取excel文件
    df = pd.read_excel(excel_file, engine='openpyxl')

    # 开始构建HTML文件
    html_content = "<html><head><title>Challenges</title>"
    html_content += "<style>.compact { margin-top: 5px; margin-bottom: 25px; }</style></head><body>"

    # 构建左侧菜单
    html_content += '<div style="float: left; width: 20%;">'
    html_content += "<ul>"
    for index, row in df.iterrows():
        html_content += f"<li><a href='#proj{index}'>{row.iloc[0] if pd.notna(row.iloc[0]) else 'N/A'} {row.iloc[5] if pd.notna(row.iloc[5]) else 'N/A'}</a></li>"
    html_content += "</ul></div>"

    # 构建详细内容
    html_content += '<div style="float: right; width: 75%;">'
    for index, row in df.iterrows():
        html_content += f"<div id='proj{index}'>"
        html_content += f"<h3>{row.iloc[0] if pd.notna(row.iloc[0]) else 'N/A'} {row.iloc[5] if pd.notna(row.iloc[5]) else 'N/A'}</h3>"
        html_content += f"<p><b>{row.iloc[1] if pd.notna(row.iloc[1]) else 'N/A'} {row.iloc[2] if pd.notna(row.iloc[2]) else 'N/A'}</b></p>"
        html_content += f"<p><a href='mailto:{row.iloc[4] if pd.notna(row.iloc[4]) else 'N/A'}'>{row.iloc[4] if pd.notna(row.iloc[4]) else 'N/A'}</a> | <a href='tel:{row.iloc[3] if pd.notna(row.iloc[3]) else 'N/A'}'>{row.iloc[3] if pd.notna(row.iloc[3]) else 'N/A'}</a></p>"

        html_content += f"<b>Limit of Participants:</b>"
        html_content += f"<p class='compact'>{row.iloc[8] if pd.notna(row.iloc[8]) else 'N/A'}</p>"

        html_content += f"<b>Would you consider continuing working on your case during the second year of the Master School?:</b>"
        html_content += f"<p class='compact'>{row.iloc[9] if pd.notna(row.iloc[9]) else 'N/A'}</p>"

        html_content += f"<b>Challenge context:</b>"
        html_content += f"<p class='compact'>{row.iloc[6] if pd.notna(row.iloc[6]) else 'N/A'}</p>"

        html_content += f"<b>Challenge description:</b>"
        html_content += f"<p class='compact'>{row.iloc[7] if pd.notna(row.iloc[7]) else 'N/A'}</p>"

        # ...[其他代码]

        html_content += "</div><hr/>"
    html_content += '</div>'

    html_content += "</body></html>"

    # 将HTML内容写入文件
    with open(html_file, 'w', encoding="utf-8") as file:
        file.write(html_content)


# 读取Excel文件并生成HTML文件
generate_html_from_excel(r"C:\Users\63166\Desktop\Cases2023.xlsx", "output.html")

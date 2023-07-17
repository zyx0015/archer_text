import re
import pandas as pd
from docx import Document
import streamlit as st

# 读取文档
def read_report(file_path):
  doc = Document(file_path)
  docx_text = [paragraph.text.strip() for paragraph in doc.paragraphs]
  docx_text = [line for line in docx_text if line.strip() != ""]
  # 删除空格和将＋替换为"与"
  docx_text = [re.sub(" ", "", line) for line in docx_text]
  docx_text = [re.sub("＋", "与", line) for line in docx_text]
  # 删除空行
  docx_text = [line for line in docx_text if not re.match(r"^\s*$", line)]
  # 打印前几行
  # 提取以图开头的行
  docx_text_number = [line for line in docx_text if re.match(r"^图.*", line)]
  docx_text_number = [line for line in docx_text_number if ":" in line]
  # 提取不以@开头的行
  docx_text_no_num = [line for line in docx_text if not re.match(r"^图.*", line)]
  return docx_text_number,docx_text_no_num


def extract_parts(lst):
      before_dot = []
      after_dot = []
      for item in lst:
          match = re.search(r'^(.*)\.(.*)$', item)
          if match:
              before_dot.append(match.group(1))
              after_dot.append(match.group(2))
      return before_dot, after_dot


#提取图号与器物名称
def artefact_name(cn_str_split):
        # 只保留数字，保留第一个罗马字符或汉字出现的时候
        index_raw, name_raw = extract_parts(cn_str_split)
        final_result = []
        final_ind = []
        # 删除空格
        for i in range(len(index_raw)):
            index_raw_n = index_raw[i]
            index_raw_n = [x for x in index_raw_n.split(".") if x.strip() != ""]
            # 定义函数来展开波浪号
            def expand_wave(x):
                parts = re.split("[～|~|-]", x)
                if len(parts) == 1:
                    return parts
                return list(range(int(parts[0]), int(parts[1])+1))
            index_n = [num for part in index_raw_n for num in expand_wave(part)]
            # 与中文和罗马数字进行匹配
            index_cn_n = name_raw[i]
            result_n = [index_cn_n] * len(index_n)
            # append
            final_result.extend(result_n)
            final_ind.extend(index_n)
        final_result = dict(zip(final_ind, final_result))
        return final_result

#提取器物编号
def artefact_number(non_cn_str_split):
    vec_b_result = []
    for item in non_cn_str_split:
        str_vec = item.split(".")
        for i in range(len(str_vec)-1, 0, -1):
            if re.match(r'^\d+$', str_vec[i]):
                prev_element = str_vec[i-1]
                if "-" in prev_element:
                    num_before_dash = re.sub(r".*?(\d+)-.*", r"\1", prev_element)
                    str_vec[i] = num_before_dash + "-" + str_vec[i]
        new_vec = str_vec.copy()
        for i in range(len(str_vec)-1, -1, -1):
            if not re.search(r"[A-Za-z\u4e00-\u9fa5]", str_vec[i]):
                j = i - 1
                while j >= 0 and not re.search(r"[A-Za-z\u4e00-\u9fa5]", str_vec[j]):
                    j -= 1
                if j >= 0:
                    prefix = re.sub(r":.*", "", str_vec[j])
                    new_prefix = prefix + ":"
                    new_vec[i] = new_prefix + str_vec[i]
        vec_b_result.extend(new_vec)
    return vec_b_result

#生成表格
def doc_to_df(docx_text_number,docx_text_no_num):
  output_table = pd.DataFrame()
  for n in range(len(docx_text_number)):
    fig_name = re.sub("\\d.*", "", docx_text_number[n])
    docx_text_number[n] = re.sub(".*?(\\d.*)", "\\1", docx_text_number[n])
    str_line = docx_text_number[n]
    # 分割字符串
    str_split = re.split("[（,）]", str_line)
    str_split = [item.strip() for item in str_split if len(item.strip()) > 0]
    # 把顿号换成点号
    str_split = [re.sub("、", ".", item) for item in str_split]
    # 分割包含中文字符和不包含中文字符的元素
    cn_str_split = [item for item in str_split if re.search("[\u4e00-\u9fa5]", item)]
    non_cn_str_split = [item for item in str_split if not re.search("[\u4e00-\u9fa5]", item)]
    if len(non_cn_str_split) != 0:
        # 对于中文的遗物名与编号向量
        final_result=artefact_name(cn_str_split)
    ####################################新分法###########################
    ##对于非中文的遗物号向量
    non_cn_str_split = [re.sub("，", ".", i) for i in non_cn_str_split]
    non_cn_str_split = [re.sub(",", ".", i) for i in non_cn_str_split]
    label_final_result=artefact_number(non_cn_str_split)
    #########################导入表格里##############################
    if len(final_result) == len(label_final_result):
        df_result = pd.DataFrame({
            'fig_name': fig_name,
            'fig_number':list(final_result.keys()),
            'artefact_name':list(final_result.values()),
            'artefact_number':label_final_result
        })
    else:
        df_result = pd.DataFrame()
    output_table = pd.concat([output_table, df_result])
  return output_table


def description_extract(output_table,label_format="（，）,"):
  label_format = st.text_input("file label format:")
  fig_des1= [f"{label_format[0]}{row['fig_name']}{label_format[1]}{row['fig_number']}{label_format[3]}" for _, row in output_table.iterrows()]
  fig_des2= [f"{label_format[0]}{row['fig_name']}{label_format[1]}{row['fig_number']}{label_format[2]}" for _, row in output_table.iterrows()]
  fig_des_all= [f"{label_format[0]}{row['fig_name']}{label_format[1]}" for _, row in output_table.iterrows()]
  
  delim_regex = "|".join([str(desc) for desc in fig_des1] + [str(desc) for desc in fig_des2])
  delim_regex_all = "|".join([str(desc) for desc in fig_des_all])
  result_des = []
  for i in range(len(fig_des1)):
      id_n = [fig_des1[i], fig_des2[i]]
      docx_text_n_1 = [text for text in docx_text_no_num if re.search(id_n[0], text)]
      docx_text_n_2 = [text for text in docx_text_no_num if re.search(id_n[1], text)]
      docx_text_n = docx_text_n_1 + docx_text_n_2
  
      split_by_fig_all = [text.split("标本") for text in docx_text_n]
      if split_by_fig_all:
          split_by_fig_all=split_by_fig_all[0]
          real_im = [split_by_fig_all[0] + split[0:] for split in split_by_fig_all if re.search(id_n[0], split) or re.search(id_n[1], split)]
          if real_im is not None:
              result_des.append(real_im)
          else:
              docx_text_n = [text for text in docx_text_no_num if re.search(fig_des_all[i], text)]
              split_by_fig_all = [text.split("标本") for text in docx_text_n]
              if split_by_fig_all:
                  split_by_fig_all=split_by_fig_all[0]
                  real_im = [split_by_fig_all[0] + split[0:] for split in split_by_fig_all if re.search(fig_des_all[i], split)]
                  if real_im is not None:
                      result_des.append(real_im)
                  else:
                      result_des.append(None)
              else:
                  result_des.append(None)
      else:
          docx_text_n = [text for text in docx_text_no_num if re.search(fig_des_all[i], text)]
          split_by_fig_all = [text.split("标本") for text in docx_text_n]
          if split_by_fig_all:
              split_by_fig_all=split_by_fig_all[0]
              real_im = [split_by_fig_all[0] + split[0:] for split in split_by_fig_all if re.search(fig_des_all[i], split)]
              if real_im is not None:
                  result_des.append(real_im)
              else:
                  result_des.append(None)
          else:
              result_des.append(None)
  result=[lst[0] if lst else None for lst in result_des]
  output_table["description"]=result
  return output_table


#执行操作
st.title("Extract description from your archaeological reports")

file = st.file_uploader("import your pdf file")
docx_text_number,docx_text_no_num=read_report(file)
output_table=doc_to_df(docx_text_number,docx_text_no_num)
final_output_table=description_extract(output_table)

# 显示数据框


# 交互式修改数据框
edited_df = st.experimental_data_editor(final_output_table,num_rows="dynamic")
st.download_button(
          label='Download output.zip',
          data=edited_df.to_csv(index=False),
          file_name='output.csv',
          mime='text/csv'
      )

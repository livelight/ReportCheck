"""
检查修订标记是否正确添加到文档中
"""

import zipfile
import xml.etree.ElementTree as ET
import re


def check_document_revisions(file_path):
    """检查文档中的修订标记"""
    print(f"\n检查文件: {file_path}")
    print("=" * 60)

    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 读取document.xml
            if 'word/document.xml' in zf.namelist():
                xml_content = zf.read('word/document.xml').decode('utf-8')
                doc_path = 'word/document.xml'
            elif 'content.xml' in zf.namelist():
                xml_content = zf.read('content.xml').decode('utf-8')
                doc_path = 'content.xml'
            else:
                print("找不到document.xml或content.xml")
                return

        # 检查修订标记
        has_del = 'w:del' in xml_content
        has_ins = 'w:ins' in xml_content

        print(f"包含 w:del (删除标记): {has_del}")
        print(f"包含 w:ins (插入标记): {has_ins}")

        # 统计修订数量
        del_count = xml_content.count('w:del')
        ins_count = xml_content.count('w:ins')
        print(f"删除标记数量: {del_count}")
        print(f"插入标记数量: {ins_count}")

        # 查找包含修订标记的段落文本
        if has_del or has_ins:
            # 使用正则表达式查找包含修订的段落
            # 匹配 <w:p ...>...</w:p> 包含 w:del 或 w:ins
            para_pattern = r'<w:p\b[^>]*>.*?</w:p>'
            paras = re.findall(para_pattern, xml_content, re.DOTALL)

            print("\n包含修订标记的段落:")
            found = 0
            for para_xml in paras:
                if 'w:del' in para_xml or 'w:ins' in para_xml:
                    # 提取所有 <w:t> 文本
                    t_pattern = r'<w:t[^>]*>([^<]*)</w:t>'
                    texts = re.findall(t_pattern, para_xml)
                    text = ''.join(texts)

                    if text.strip():
                        found += 1
                        print(f"\n  段落 {found}: {text[:100]}...")
                        if 'w:del' in para_xml:
                            print(f"    [OK] 包含删除标记 (w:del)")
                        if 'w:ins' in para_xml:
                            print(f"    [OK] 包含插入标记 (w:ins)")

            if found == 0:
                print("  (找到修订标记但无法提取段落文本，请手动检查文档)")

    except Exception as e:
        print(f"检查失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("=" * 60)
    print("修订标记检查结果")
    print("=" * 60)

    # 检查修订后的文档
    check_document_revisions('test_sample_revised_tc.wps')
    check_document_revisions('test_sample_revised_tc.wpsx')
    check_document_revisions('test_sample_revised_track_changes.docx')

    print("\n" + "=" * 60)
    print("结论:")
    print("  如果 w:del 和 w:ins 都显示为 True，")
    print("  则修订标记已正确添加到文档中。")
    print("  可在WPS/Word的'审阅'选项卡中查看修订。")
    print("=" * 60)

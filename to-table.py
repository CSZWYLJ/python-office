from docx import Document


def hcf(*x):                                #计算最大公约数
    smaller=min(x)
    for i in reversed(range(1,smaller+1)):
        if list(filter(lambda j: j%i!=0,x)) == []:
            return i

class WordOutlineGenTable():
    def __init__(self, source_docx):
        self.source_docx = source_docx

    def analyse_outline(self):
        source_doc = Document(self.source_docx)
        paragraph_list = source_doc.paragraphs

        wrap_dict = {}
        normal_wrap_num = 0
        parent_wrap_level = None
        # 保存标题等级
        header_list = []

        for paragraph in paragraph_list:
            header_level = paragraph.style.name
            print(header_level, paragraph.text)
            caption_wrap = self.CellWrap(paragraph.text)
            caption_wrap.level = header_level
            caption_wrap_list = []
            if header_level not in wrap_dict:
                content_count = 0
                if header_level is "Normal" and paragraph.text.startswith("- "):
                    content_count += 1
                # 不存在字典中时，先创建list，添加对象到list，更新这个键值对
                caption_wrap_list.append(caption_wrap)
                wrap_dict[header_level] = caption_wrap_list
                header_list.append(header_level)
            else:
                # 已经存在一个标题时，统计单元格数量
                if header_level is "Normal" and len(wrap_dict[header_level]) == 1:
                    # 保存正文和标题第一次出现的位置
                    parent_wrap_level = header_list[-1]
                    normal_wrap_num+=1
                # 存在于字典中时，需要向值为list类型的对象中append对象
                wrap_dict[header_level].append(caption_wrap)
        return wrap_dict

    def get_wrap_chunk(self):
        pass

    def word_engine(self):
        pass

    class CellWrap():
        def __init__(self, content, merge_num=None, level=None):
            # 大纲解析出的单元格内容
            self.content = content
            # 大纲解析出的单元格等级，
            # 规定：1. 对标题等级来说，必须为不同等级，即1级后面必须是2级，2级后面必须是3级，否则抛出异常
            #      2. 对正文内容来说，同级段落纵向写入，段落里面的具体条例必须横向写入
            #  否则直接抛出解析异常
            self.level = level
            # 等级的数量，例如1级标题占用20个纵向单元格
            self.merge_num = merge_num
            # 上一级是谁
            self.parent_cell = None
            # 下一级节点
            self.next_cell = []
            # 同级节点的list
            self.same_level_cell = []
            # 判断是否纵向合并，一般标题默认True纵向，正文False横向
            self.portrait_flag = True


if __name__ == "__main__":
    print(hcf(10,25,35,65))

    word_table = WordOutlineGenTable("./res-JKB.docx")
    word_table.analyse_outline()
    pass

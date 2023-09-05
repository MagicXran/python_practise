try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET


def main():
    file = r"C:\Repository\Nercar\ShaGang\RM改造\Materials\L1\ModelComm.xml"
    tree = (parse_xml(file))

    print(tree.tag)  # 打印root节点元素
    for elem in tree.iter():
        print(elem.tag, elem.attrib)
        # for attr in elem.attrib.values():
        #     print(attr)
        # for sub_el in elem.iter():
        #     print(sub_el)
        # print(elem.tag, elem.attrib)
    #     只遍历那些有指定标签的元素  # for elem in tree.iter(tag='Module'):  # print(elem.tag, elem.attrib)

    # for elem in tree.iterfind('*/Module[@Name="RM_SET"]/*/[@Name]'):  # for elem in tree.iterfind('[tag="Module"]'):  #     print(elem.tag, elem.attrib)  #     print(elem.get('Name'))


def parse_xml(file):
    return ET.fromstring(xml2str(file))
    pass


def xml2str(file):
    """
    将xml文件读取转为utf-8格式存于内存
    :param file:
    :return:
    """
    with open(file, encoding="utf-8") as f:
        # return gbk2utf8(f.read())
        return f.read()
    pass


def gbk2utf8(xml_str: str):
    """
    将 xml中编码改为utf-8
    :param xml_str:
    :return:
    """
    if xml_str.find('GBK'):
        temp_txt = xml_str.replace("utf-8", 'GBK')
        return temp_txt
    else:
        return xml_str
    pass


if __name__ == '__main__':
    # main()  # test()
    tree = ET.ElementTree(file=r"C:\Repository\Nercar\ShaGang\RM改造\Materials\L1\ModelComm.xml")
    root = tree.getroot()
    print(root.tag, root.attrib)

    count_var = 0
    # 每个节点均包含节点名tag和属性attrib
    for Links in root:
        print(Links.tag, Links.attrib)
        for Link in Links:
            print(Link.tag, Link.attrib)
            for Properties in Link:
                # print(Properties.tag, Properties.attrib)
                if Properties.tag == 'SendMessages':
                    for SendMessages in Properties:
                        print('SendMessages:', SendMessages.tag, SendMessages.attrib)
                        if SendMessages.attrib.get('Type') == '202':
                            for prop in SendMessages:

                                for vars in prop:
                                    # print('--', vars.tag, vars.attrib)
                                    if vars.tag == 'Variable':
                                        print(vars.tag, vars.attrib)
                                        count_var = count_var + 1
                            # for Property in prop:
                            #     print(Property.tag,Property.attrib)

                            # if Property.tag == 'Property' and Property.attrib.get('Value') =='MDS_L2_ALIVE':
                            #     print(Property.tag, Property.attrib)

                            # print(Property.attrib.get('Name'))

                    pass

    print('count',count_var)
    pass

from spider import *
import sys

if __name__ == '__main__':
    spider_list = (
        (RankSpider, '排名'),
        (SiteSpider, '收录'),
        (CheckSpider, '核对')
    )
    engine_list = (
        (SMRuler, '神马'),
        (SogouPCRuler, '搜狗PC'),
        (SogouMobileRuler, '搜狗MOBILE'),
        (BaiduPCRuler, '百度PC'),
        (BaiduMobileRuler, '百度MOBILE'),
        (SLLPCRuler, '360PC'),
        (SLLMobileRuler, '360MOBILE'),
    )

    if len(sys.argv) >= 2:
        spider_index = sys.argv[1]
    else:
        spider_index = input('''要查找什么数据？
%s
''' % '\n'.join(['%s 请输入：%s' % (ruler_name, i) for (i, (_, ruler_name)) in enumerate(spider_list)]))

    if len(sys.argv) >= 3:
        engine_index = sys.argv[2]
    else:
        engine_index = input('''要查找哪个搜索引擎？
%s
''' % '\n'.join(['%s 请输入：%s' % (ruler_name, i) for (i, (_, ruler_name)) in enumerate(engine_list)]))

    (spider_class, spider_name) = spider_list[int(spider_index)]
    (ruler_class_, ruler_name) = engine_list[int(engine_index)]
    os.system('title %s%s' % (ruler_name, spider_name))
    spider_class(ruler_class_)

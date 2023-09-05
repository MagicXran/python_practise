# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html

import codecs
import json

# useful for handling different item types with a single interface
from scrapy.pipelines.images import ImagesPipeline


class ArticlespiderPipeline:
    # item 即 jobbole中yield的article_items
    def process_item(self, item, spider):
        return item


class JsonWithEncodingPipeline(object):
    """
    自定义json文件到处,即将文件保存到本地
    """

    def __init__(self):
        self.file = codecs.open('article.json', 'a', encoding='utf-8')

    # override方法
    def process_item(self, item, spider):
        # 将对象转化为json格式存储(str类型)
        lines = json.dumps(dict(item), ensure_ascii=False) + "\n"
        self.file.write(lines)
        return item

    def spider_close(self, spider):
        self.file.close()


class ArticleImagePipeline(ImagesPipeline):
    """
    修改下载的图片路径
    """

    # item 即 jobbole中yield的article_items
    # override method 进行数据拦截
    def item_completed(self, results, item, info):
        if "front_image_url" in item:
            for ok, value in results:
                image_file_path = value["path"]
            item['front_image_url'] = image_file_path

        # 由于在setting.py设置了pipeline的优先级, 所以此自定义执行完后跳转到 ArticlespiderPipeline 开始执行.
        return item

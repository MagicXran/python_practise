# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class ArticlespiderItem(scrapy.Item):
    """
    Item: 是一个高级的map,即dict
        所有字段类型都是 Field, 没有int, str ... 之分
    """
    # define the fields for your item here like:
    # name = scrapy.Field()
    # create_date = scrapy.Field()
    # url = scrapy.Field()
    # url_obj_id = scrapy.Field()
    front_image_url = scrapy.Field()
    # front_img_path = scrapy.Field()
    # praise_num = scrapy.Field()
    # comments_num = scrapy.Field()

    pass

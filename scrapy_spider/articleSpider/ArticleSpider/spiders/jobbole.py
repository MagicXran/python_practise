from urllib import parse

import scrapy
from scrapy import Request

from articleSpider.ArticleSpider.items import ArticlespiderItem


class JobboleSpider(scrapy.Spider):
    """
    该爬虫作用:
    1. 获取博客园首页所有标题,点赞,评论数,发布时间以及多少人浏览
    """
    name = 'jobbole'
    allowed_domains = ['news.cnblogs.com']
    start_urls = ['http://news.cnblogs.com/']

    def parse(self, response):
        """
        0. parse目的:
        1. 获取当前新闻预览页所有资料的url,交给scrapy进行下载后调用相应的解析方法.
        2. 获取下一页的url,交给scrapy进行下载,后交给parse继续跟进.
        :param response:
        :return:
        """

        #### 1

        # 第一种: 返回列表,要判断列表是否为空
        # if response.xpath(r'//a[@href]').extract():
        #     url = response.xpath(r'//a[@href]').extract()

        # 第二种方法: 返回列表中第一个元素,如果为空则返回""
        # url = response.xpath(r'//a[@href]').extract_first("")

        # 第三种方法: 使用css选择器
        # 获取所有 class为news_entry的 子元素为a的其href属性值.
        url = response.css(".news_entry a::attr(href)").extract()

        # 获取首页中所有的新闻块(含标题,链接,推荐数,图片,大纲,投递人,评论数,浏览数,tag,发布时间等)
        # post_nodes 是一个 SelectorList  , 注意不能使用 extract(), 因为如此会变成str类型.
        post_nodes = response.xpath("//div[@class='news_block']")[:2]
        # post_nodes = response.css(".news_block")

        # post_node 是一个 Selector
        # Selector 可以继续调用css选择器
        for post_node in post_nodes:
            # 获取每个链接的附带的图片, 位于:class为entry_summary的子元素为a的子元素为img的src属性中.
            image_url = post_node.css('.entry_summary a img::attr(src)').extract_first('')
            post_url = post_node.css('.news_entry a::attr(href)').extract_first('')

            # 每一条都交给scrapy,马上处理.
            # parse.urljoin(response.url, post_url): 当 post_url 为相对路径时(/n/706905/),自动加上response.url根路径(
            # https://news.cnblogs.com/), 如果 post_url 是绝对路径时(https://news.cnblogs.com/n/706905/), response.url不会加入.
            #
            # 注意: 这个 Request 对应的下载,不是当前这个parse(),Request对应的是详情界面
            #   所以需要自定义个一个函数,去 parse 该 Request 的资源.
            yield Request(url=parse.urljoin(response.url, post_url), meta={"front_image_url": image_url},
                          callback=self.parse_details, dont_filter=False)  # dont_filter: 设置已下载过的url进行过滤,但会内容更新会错过.
            # callback 默认回调 parse(self, response)
            # 此中可以看出 => scrapy 采用的是深度优先策略,遇到url就交给scrapy engine 进行下载, 然后再进一步解析url.
            # 单步调试时发现, Request 后又直接开始新的一次循环,而非马上跳转到 parse_details() ,
            # 其原因是: scrapy是异步io框架(使用单线程即可实现高并发)即 Request 后会交给 parse_details(), 然后就会继续执行,此时服务器还未返回数据,所以就进行进一步的循环.

        #### 2
        # 提取下一页,并交给scrapy进行下载
        # 两种办法提却下一页url:
        # a. 使用css, 获取a,判断其值为`Next >`( <a ..>Next > </a> ) 即是下一页的url
        #      获取 class=pager 的子元素的最后一个 a 的值.

        # next_url = response.css('div.pager a:last-child::text').extract_first("")
        # if next_url == 'Next >':
        #     #  确定是下一页,则提却其 href 属性值
        #     next_url = response.css('div.pager a:last-child::attr(href)').extract_first("")
        #     yield Request(parse.urljoin(response.url, next_url), callback=self.parse)  # 即使不显式写callback,也会默认调用 parse()

        # b. 使用 xpath 在此情形中更加方便.
        #   搜索所有 a 标签,且满足: text 的值中包含文本 'Next >', 但是得到的是一个完整的 a 标签. 我们只需要其中的 href 值.
        #  所以在最后添加 /@href 获取其中url.

        # next_url = response.xpath("//a[contains(text(),''Next >'')]/@href").extract_first('')
        # yield Request(parse.urljoin(response.url, next_url))  # 即使不显式写callback,也会默认调用 parse()
        pass

    def parse_details(self, response):
        title = response.xpath('//div[@class="secondary-text text-center ng-tns-c140-11"]/a/text()')
        article_item = ArticlespiderItem()

        # 获取response中meta的值,存入items中. 其中meta的值是  parse() 传来的.
        if response.meta.get("front_image_url", ''):
            article_item['front_image_url'] = [response.meta.get("front_image_url", '')]
        else:
            article_item['front_image_url'] = []
        print(article_item["front_image_url"])

        # 进入 pipeline
        yield article_item
        # 与字典赋值一样
        # get(name,default) 对应的name有值则返回,为空则默认给""

        # article_item["front_image_url"] = response.meta.get("front_image_url", "")
        # article_item["url"] = response.url
        # article_item["url_obj_id"] = common.get_md5(response.url)
        # pass

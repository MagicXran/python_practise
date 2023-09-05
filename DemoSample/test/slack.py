import logging

from slack_logger import SlackHandler, SlackFormatter

sh = SlackHandler('https://hooks.slack.com/...')  # url is like 'https://hooks.slack.com/...'
sh.setFormatter(SlackFormatter())
logging.basicConfig(handlers=[sh])
logging.warning('warn message')

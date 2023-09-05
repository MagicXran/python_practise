import sys

import pygame

# 1.初始化操作
pygame.init()

# 2.创建游戏窗口
window_size_width = 600
window_size_height = 400
window = pygame.display.set_mode((window_size_width, window_size_height))
window.fill((199, 198, 182))
# 设置游戏标题
pygame.display.set_caption('课表导出软件  for my lover : 小鱼')

font = pygame.font.Font("C:/Windows/Fonts/STXINWEI.TTF", 30)

# 1.确定按钮
bx1, by1, bw1, bh1 = 30, 100, 100, 50
pygame.draw.rect(window, (255, 255, 255), (bx1, by1, bw1, bh1))
text1 = font.render("确认", True, (239, 246, 252))
tw1, th1 = text1.get_size()
tx1, ty1 = bx1 + bw1 / 2 - tw1 / 2, by1 + bh1 / 2 - th1 / 2
# 将准备好的文本信息，绘制到主屏幕 Screen 上。
window.blit(text1, (tx1, ty1))

# 2.取消按钮
bx2, by2, bw2, bh2 = 30, 200, 100, 50
pygame.draw.rect(window, (0, 255, 0), (bx2, by2, bw2, bh2))
text2 = font.render("取消", True, (255, 255, 255))
tw2, th2 = text2.get_size()
tx2, ty2 = bx2 + bw2 / 2 - tw2 / 2, by2 + bh2 / 2 - th2 / 2
window.blit(text2, (tx2, ty2))

pygame.display.update()

# 3.让游戏保持一直运行的状态
while True:
    # 4.检测事件
    for event in pygame.event.get():
        # 对事件作出相应的响应
        if event.type == pygame.QUIT:  # 如果点击了关闭按钮
            sys.exit()

        if event.type == pygame.MOUSEBUTTONDOWN:  # 如果鼠标按下
            mx, my = event.pos  # 获取鼠标点击的位置
            if bx1 + bw1 >= mx >= bx1 and by1 + bh1 >= my >= by1:
                pygame.draw.rect(window, (200, 200, 200), (bx1, by1, bw1, bh1))
                window.blit(text1, (tx1, ty1))
                pygame.display.update()
                print("确定按钮被点击")
            elif bx2 + bw2 >= mx >= bx2 and by2 + bh2 >= my >= by2:
                pygame.draw.rect(window, (200, 200, 200), (bx2, by2, bw2, bh2))
                window.blit(text2, (tx2, ty2))
                pygame.display.update()
                print("取消按钮被点击")

        if event.type == pygame.MOUSEBUTTONUP:
            mx, my = event.pos  # 获取鼠标点击的位置
            if bx1 + bw1 >= mx >= bx1 and by1 + bh1 >= my >= by1:
                pygame.draw.rect(window, (255, 0, 0), (bx1, by1, bw1, bh1))
                window.blit(text1, (tx1, ty1))
                pygame.display.update()
                print("确定按钮被松开")
            elif bx2 + bw2 >= mx >= bx2 and by2 + bh2 >= my >= by2:
                pygame.draw.rect(window, (0, 255, 0), (bx2, by2, bw2, bh2))
                window.blit(text2, (tx2, ty2))
                pygame.display.update()
                print("取消按钮被松开")

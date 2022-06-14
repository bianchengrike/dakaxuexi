from pywinauto.application import Application



# 类名连接word，不知道为啥换了办公室电脑，所有控件都好找了，就如此简单？但是一会又不好找了！？
# app = Application(backend='uia').connect(process=process_id)
app = Application(backend='uia').connect(class_name='OpusApp')

# 研究了一下，office2013-word都有这第一个页面
# 找到视图菜单项并点击
# .print_control_identifiers() 这个非常重要！执行完了要删掉，不然下面找不到
# win = app.window(title_re='.* - Word')
win = app.window(class_name='OpusApp')
win.set_focus()
crl=win.window(title='阅读版式菜单(&F)')
crl.wait('ready')
crl.MenuItem2.click_input()
# 找到编辑文档并点击进入到第二个页面
crl2= win.window(title='编辑文档')
crl2.wait('ready')
crl2.click_input()

# 第二个页面，都优化了一下
# 找到审阅并选择
crl3=win.window(title='功能区选项卡')
crl3.wait('ready')
crl3.select('审阅') 
# 找到限制编辑并点击
crl4=win.window(title="限制编辑")
crl4.wait('ready')
crl4.click_input()
# 找到停止保护按钮并点击
crl5=win.window(title="停止保护")
crl5.wait('ready')
crl5.click_input()
# 找到取消保护窗口的编辑框并输入密码123+回车
crl6=win.window(title='取消保护文档')
crl6.wait('ready')
crl6.Edit.type_keys('123'+'{ENTER}')

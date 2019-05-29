# 导入模块
import os
import time

import xlrd
import xlwt
from wxpy import *


class bot_groups:
    def write_sheetRow(self, sheet, row_value_list, row_index, is_bold):
        i = 0
        style = xlwt.easyxf('font: bold 1')
        for svalue in row_value_list:
            if is_bold:
                sheet.write(row_index, i, svalue, style)
            else:
                sheet.write(row_index, i, svalue)
            i = i + 1

    def make_send_excel(self, cur_groups, config_group_names):
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet('sheet1', cell_overwrite_ok=True)
        head_list = ['puid', 'name', 'is_send']

        row_index = 0
        self.write_sheetRow(sheet, head_list, row_index, True)

        config_count = 0
        for group in cur_groups:
            row_index = row_index + 1
            is_in_config = 1 if group.name in config_group_names else 0
            value_list = [group.puid, group.name, is_in_config]
            self.write_sheetRow(sheet, value_list, row_index, False)

            if is_in_config:
                config_count += 1
        file_name = os.path.join(os.getcwd(), 'send.xlsx')
        wbk.save(file_name)
        print('生成群列表时，当前群匹配到配置文件中的[%s]个群' % config_count)

    def read_send_excel(selfs):
        readbook = xlrd.open_workbook(r'send.xlsx')
        sheet = readbook.sheet_by_index(0)
        rows = sheet.nrows
        datas = []
        for row in range(1, rows):
            data = {
                'puid': sheet.cell(row, 0),
                'name': sheet.cell(row, 1),
                'is_send': sheet.cell(row, 2)
            }
            if data.get('is_send').value == 1:
                datas.append(data)
        return datas

    def read_config_groups(self):
        read_book = xlrd.open_workbook(r'config_test.xlsx')
        sheet = read_book.sheet_by_index(0)
        rows = sheet.nrows
        datas = []
        for row in range(0, rows):
            data = sheet.cell(row, 0).value
            datas.append(data)
        return datas


if __name__ == '__main__':
    # 初始化机器人，扫码登陆
    bot = Bot(cache_path=False)
    # bot = Bot()
    bot.enable_puid()

    b = bot_groups()

    # 获取配置的要发送的群名称
    print('正在获取配置文件中要发送的群...')
    config_group_names = b.read_config_groups()
    print('完成获取配置文件中要发送的群.共[%s]个群' % len(config_group_names))

    print('开始获取当前微信号中的群...')
    cur_groups = bot.groups()
    print('完成获取当前微信号中的群.共[%s]个群' % len(cur_groups))

    # 保存excel
    print('开始生成要发送的群列表...')
    b.make_send_excel(cur_groups, config_group_names)
    print('完成生成要发送的群列表')

    # 监听文件助手
    print('开始监听文件助手...')
    file_helper = bot.file_helper


    @bot.register(file_helper, except_self=False)
    def forward_boss_message(msg):
        print('开始获取已保存待发送的群...')
        send_groups = b.read_send_excel()
        print('完成获取已保存待发送的群.共[%s]个群' % len(send_groups))
        i = 0
        for g in send_groups:
            puid = g.get('puid').value
            s_groups = cur_groups.search(puid=puid)
            s_group = s_groups[0]
            i = i + 1
            msg.forward(s_group)
            print('已发送到[%s](%s/%s)' % (s_group.nick_name, i, len(send_groups)))
            time.sleep(3)
        print('发送完成')


    print('完成监听文件助手')

    # bot.join()
    embed()

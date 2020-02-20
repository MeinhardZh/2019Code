import wx
import sys
import os
import wx.adv
import wx.grid
import BE
import datetime

APP_TITLE = '疫情防控'
columns_name = ['乡镇街道', '今日累计居家观察人数', '今日新增人数', '今日累计解除',
                '今日解除', '现有居家观察人数', '昨日累计居家观察人数', '昨日累计解除人数', '昨日正在居家观察人数']


class mainFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, -1, APP_TITLE)

        self.SetBackgroundColour(wx.Colour(224, 224, 224))
        self.SetSize((1000, 500))
        self.Center()
        
        self.grid = self.CreateGrid(self)
        self.search_data = []

        now = (datetime.datetime.now() + datetime.timedelta(days=1))
        self.seldateList = [now.year, now.month, now.day - 1]
        self.seldate = datetime.date(self.seldateList[0], self.seldateList[1], self.seldateList[2])

        btn_insert = wx.Button(self, -1, '导入数据', pos=(20, 10), size=(130, 30))
        btn_export = wx.Button(self, -1, '生成表格', pos=(170, 10), size=(130, 30))
        select_icon = wx.adv.DatePickerCtrl(
            self, -1, pos=(320, 10), size=(130, 30), style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        btn_search = wx.Button(self, -1, '查询数据', pos=(470, 10), size=(130, 30))
        btn_close = wx.Button(self, -1, '关闭系统', pos=(620, 10), size=(130, 30))

        # 控件事件
        self.Bind(wx.EVT_BUTTON, self.OnInsert, btn_insert)
        self.Bind(wx.EVT_BUTTON, self.OnExport, btn_export)
        self.Bind(wx.EVT_BUTTON, self.OnSearch, btn_search)
        self.Bind(wx.EVT_BUTTON, self.OnClose, btn_close)

        self.Bind(wx.adv.EVT_DATE_CHANGED, self.OnCalSelChanged, select_icon)

        # 系统事件
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        # 连接数据库
        self.business = BE.Business()

    def OnClose(self, evt):

        mesdlg = wx.MessageDialog(None, '确定关闭本系统？', '操作提示',
                                  wx.YES_NO | wx.ICON_QUESTION)
        if mesdlg.ShowModal() == wx.ID_YES:
            self.Destroy()

    def OnExport(self, evt):
        if len(self.search_data) != 0:
            dialog = wx.DirDialog(self, "选择文件夹", style=wx.DD_DEFAULT_STYLE)
            if dialog.ShowModal() == wx.ID_OK:
                filepath = dialog.GetPath()
                try:
                    self.business.export_data(
                        self.search_data, filepath, self.seldate)
                except:
                    mesdlg = wx.MessageDialog(
                        None, "导出失败! 请联系管理员", '', wx.OK | wx.ICON_ERROR)
                    dialog.Destroy
                    if mesdlg.ShowModal() == wx.ID_OK:
                        mesdlg.Destroy
                else:
                    mesdlg = wx.MessageDialog(
                        None, "导出成功！", '', wx.OK | wx.ICON_INFORMATION)
                    dialog.Destroy()
                    if mesdlg.ShowModal() == wx.ID_OK:
                        mesdlg.Destroy

        else:
            mesdlg = wx.MessageDialog(
                None, "未查询数据！", '', wx.OK | wx.ICON_ERROR)
            if mesdlg.ShowModal() == wx.ID_OK:
                mesdlg.Destroy

    def OnSearch(self, evt):
        try:
            self.search_data = self.business.search_data(self.seldate)
        except:
            mesdlg = wx.MessageDialog(
                None, "查询失败！请联系管理员", '', wx.OK | wx.ICON_ERROR)
            if mesdlg.ShowModal() == wx.ID_OK:
                mesdlg.Destroy
        else:
            try:
                self.UpdateGrid()
            except:
                mesdlg = wx.MessageDialog(
                    None, "更新表失败！请联系管理员", '', wx.OK | wx.ICON_ERROR)
                if mesdlg.ShowModal() == wx.ID_OK:
                    mesdlg.Destroy
            else:
                mesdlg = wx.MessageDialog(
                        None, "查询成功！", '', wx.OK | wx.ICON_INFORMATION)
                if mesdlg.ShowModal() == wx.ID_OK:
                    mesdlg.Destroy

    def OnInsert(self, evt):
        wildcard = 'All files(*.*)|*.*'
        dialog = wx.FileDialog(
            None, '选择一个文件', os.getcwd(), '', wildcard, wx.FD_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            filename = dialog.GetPath()
            try:
                self.business.import_data(filename)
            except:
                mesdlg = wx.MessageDialog(
                    None, "导入失败! 请联系管理员", '', wx.OK | wx.ICON_ERROR)
                dialog.Destroy
                if mesdlg.ShowModal() == wx.ID_OK:
                    mesdlg.Destroy
            else:
                mesdlg = wx.MessageDialog(
                    None, "导入成功！", '', wx.OK | wx.ICON_INFORMATION)
                dialog.Destroy
                if mesdlg.ShowModal() == wx.ID_OK:
                    mesdlg.Destroy

    def OnCalSelChanged(self, evt):
        cal = evt.GetEventObject()
        datestr = cal.GetValue()

        self.seldateList[0] = datestr.year
        self.seldateList[1] = datestr.month + 1
        self.seldateList[2] = datestr.day
        self.seldate = datetime.date(self.seldateList[0], self.seldateList[1], self.seldateList[2])

    def CreateGrid(self, parent):
        grid = wx.grid.Grid(parent, pos=(0, 70))

        grid.CreateGrid(1, len(columns_name))

        for i in range(len(columns_name)):
            grid.SetColLabelValue(i, columns_name[i])
            grid.SetReadOnly(0, i, isReadOnly=True)

        grid.AutoSize()

        return grid

    def UpdateGrid(self):

        self.grid.SetCellValue(0, 0, '双土')
        self.grid.SetCellValue(
            0, 1, str(self.search_data.get('today_total_observe')))
        self.grid.SetCellValue(0, 2, str(self.search_data.get('today_add')))
        self.grid.SetCellValue(
            0, 3, str(self.search_data.get('today_total_relieve')))
        self.grid.SetCellValue(
            0, 4, str(self.search_data.get('today_relieve')))
        self.grid.SetCellValue(
            0, 5, str(self.search_data.get('today_still_observe')))
        self.grid.SetCellValue(
            0, 6, str(self.search_data.get('yesterday_total_observe')))
        self.grid.SetCellValue(
            0, 7, str(self.search_data.get('yesterday_total_relieve')))
        self.grid.SetCellValue(
            0, 8, str(self.search_data.get('yesterday_still_observe')))


class mainApp(wx.App):
    def OnInit(self):
        self.SetAppName(APP_TITLE)
        self.Frame = mainFrame(None)
        self.Frame.Show()
        # self.Frame.Fit()
        return True


if __name__ == "__main__":
    app = mainApp()
    app.MainLoop()

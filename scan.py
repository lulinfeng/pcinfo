# coding:utf-8

import wmi, platform, time
from os import path
import win32api, win32com.client
import wx
import wx.lib.gizmos as gizmos


class main_frame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self, None, -1, 'PC数据收集测试版',
			size=(500, 400),
			style=wx.DEFAULT_FRAME_STYLE ^ wx.MAXIMIZE_BOX ^ wx.RESIZE_BORDER,
		)
		statusbar = self.CreateStatusBar(2)
		statusbar.SetStatusWidths([-6, -3])
		statusbar.SetStatusText('2020.03.15   811191000@qq.com', 0)
		statusbar.SetStatusText("作者：卢琳峰   build 1.0", 1)
		panel=wx.Panel(self, -1)
		led = gizmos.LEDNumberCtrl(panel, -1, (25,175), (280, 50), gizmos.LED_ALIGN_CENTER)
		self.clock = led
		self.timer = wx.Timer(self)
		self.timer.Start(1000)
		self.OnTimer(None)

		self.Bind(wx.EVT_TIMER, self.OnTimer)


		self.name = wx.TextCtrl(panel, -1, size=(100, -1))
		name = wx.StaticText(panel, -1, '姓名 :')
		# self.model = wx.ComboBox(panel, -1, '台式机', choices=['台式机', '笔记本'])
		self.model = '台式机'

		dept_choices = []
		with open('./dept.ini', 'r', encoding='utf-8') as f:
			dept_choices = [i.strip() for i in f.readlines() if i.strip()]
		self.dept = wx.ComboBox(
			panel, -1,
			dept_choices[0] if dept_choices else '',
			size=(100, -1),
			choices=dept_choices)

		self.other = wx.TextCtrl(panel, -1, size=(180, -1))
		self.mem_size = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180,-1))
		self.disk = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180,-1))
		self.cpu = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180, -1))
		self.board = wx.TextCtrl(panel,-1, size=(180,-1))
		self.monitor = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180, -1))
		self.display = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180, -1))
		self.ip = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180, -1))
		self.mac = wx.TextCtrl(panel, -1, style=wx.TE_MULTILINE, size=(180, -1))

		button=wx.Button(panel,-1,'检测')
		buttons=wx.Button(panel,-1,'保存')
		self.Bind(wx.EVT_BUTTON, self.start, button)
		self.Bind(wx.EVT_BUTTON, self.toexl, buttons)

		mainSizer=wx.BoxSizer(wx.VERTICAL)
		gbs = wx.GridBagSizer(vgap=5, hgap=5)

		gbs.Add(name, (1, 0))
		gbs.Add(self.name, (1, 1))
		size_h = wx.BoxSizer(wx.HORIZONTAL)
		pc_btn = wx.RadioButton(panel, -1, '台式机', style=wx.RB_GROUP)
		pc_btn2 = wx.RadioButton(panel, -1, '笔记本' )
		pc_btn.Bind(wx.EVT_RADIOBUTTON, self.on_btn)
		pc_btn2.Bind(wx.EVT_RADIOBUTTON, self.on_btn)

		size_h.Add(pc_btn)
		size_h.Add(pc_btn2)
		gbs.Add(size_h, (1,3))
		bm = wx.StaticText(panel, -1, '部门:')
		gbs.Add(bm, (2, 0))
		gbs.Add(self.dept, (2, 1))

		other=wx.StaticText(panel, -1, '备注:')
		gbs.Add(other, (2, 2))
		gbs.Add(self.other, (2, 3))

		board = wx.StaticText(panel, -1, '主板 :')
		gbs.Add(board, (3, 0))
		gbs.Add(self.board, (3, 1))

		monitor = wx.StaticText(panel, -1, '显示器:')
		gbs.Add(monitor, (3, 2))
		gbs.Add(self.monitor, (3, 3))

		cpu = wx.StaticText(panel, -1, 'CPU :')
		gbs.Add(cpu, (4, 0))
		gbs.Add(self.cpu, (4, 1))
		display = wx.StaticText(panel, -1, '显卡 :')
		gbs.Add(display, (4, 2))
		gbs.Add(self.display, (4, 3))

		disk = wx.StaticText(panel, -1, '硬盘 :')
		gbs.Add(disk, (5, 0))
		gbs.Add(self.disk, (5, 1))
		memsize=wx.StaticText(panel, -1, '内存 :')
		gbs.Add(memsize, (5, 2))
		gbs.Add(self.mem_size, (5, 3))

		ip = wx.StaticText(panel, -1, 'I P :')
		gbs.Add(ip, (6, 0))
		gbs.Add(self.ip, (6, 1))
		mac = wx.StaticText(panel, -1, 'MAC :')
		gbs.Add(mac, (6, 2))
		gbs.Add(self.mac, (6, 3))

		mainSizer.Add(led,0,wx.EXPAND|wx.TOP|wx.DOWN,5)
		mainSizer.Add(gbs, 1, wx.EXPAND | wx.ALL, 10)

		bsizer=wx.BoxSizer(wx.HORIZONTAL)
		bsizer.Add(button,0,wx.ALIGN_CENTER)
		bsizer.Add((20,20))
		bsizer.Add(buttons)
		mainSizer.Add(bsizer,0,wx.DOWN|wx.TOP|wx.ALIGN_CENTER,10)

		panel.SetSizer(mainSizer)
		mainSizer.Fit(self)
		mainSizer.SetSizeHints(self)

	def getFileVersion(self, file_name):
		info = win32api.GetFileVersionInfo(file_name,'\\')
		ms = info['FileVersionMS']
		ls = info['FileVersionLS']
		version = '%d.%d.%d.%04d' % (win32api.HIWORD(ms), win32api.LOWORD(ms), win32api.HIWORD(ls), win32api.LOWORD(ls))
		return version

	def OnTimer(self, evt):
		t = time.localtime(time.time())
		st = time.strftime("%Y   %H:%M:%S", t)
		self.clock.SetValue(st)

	def start(self,evt):
		for i in (self.other,
			self.board, self.monitor,
			self.cpu, self.display,
			self.disk, self.mem_size,
			self.ip, self.mac,
		):
			i.Clear()
		w=wmi.WMI()
		# file_name=environ['programfiles']+r'\Internet Explorer\iexplore.exe'
		# self.ie_ver.AppendText(self.getFileVersion(file_name))
		self.other.AppendText(platform.platform())
		#操作系统
		#w.Win32_SystemOperatingSystem()[0].PartComponent
		# https://stackoverflow.com/questions/14227171/how-to-get-memory-information-ram-type-e-g-ddr-ddr2-ddr3-with-wmi-c
		_mem_type = {
			0x01: 'Other',
			0x02: 'Unknown',
			0x03: 'DRAM',
			0x04: 'EDRAM',
			0x05: 'VRAM',
			0x06: 'SRAM',
			0x07: 'RAM',
			0x08: 'ROM',
			0x09: 'FLASH',
			0x0A: 'EEPROM',
			0x0B: 'FEPROM',
			0x0C: 'EPROM',
			0x0D: 'CDRAM',
			0x0E: '3DRAM',
			0x0F: 'SDRAM',
			0x10: 'SGRAM',
			0x11: 'RDRAM',
			0x12: 'DDR',
			0x13: 'DDR2',
			0x14: 'DDR2 FB-DIMM',
			0x15: 'Reserved',
			0x16: 'Reserved',
			0x17: 'Reserved',
			0x18: 'DDR3',
			0x19: 'FBD2',
			0x1A: 'DDR4',
			0x1B: 'LPDDR',
			0x1C: 'LPDDR2',
			0x1D: 'LPDDR3',
			0x1E: 'LPDDR4',
		}

		#内存
		mem_info = []
		for i in w.win32_physicalmemory():
			mem_info.append('%s: %s %dGB' % (
				i.DeviceLocator,
				_mem_type.get(i.SMBIOSMemoryType, 'Other'),
				int(i.capacity) / (1024 ** 3))
			)
			# self.mem_size.AppendText('%s:%dM\n'%(i.tag,int(i.capacity)/(1024*1024)))
		self.mem_size.AppendText('\n'.join(mem_info))

		# for i in w.win32_computersystem():
		# 	print(i)
		# 	self.mem_size.AppendText(u'可用物理内存%.2fM'%(int(i.TotalPhysicalMemory)/1024.0/1024))
		# 	#加入域 domainrole返回1
		# 	# print(i)
		# 	self.domain.AppendText('%s %s' %({0:u'独立工作站',1:u'成员工作站',2:u'独立服务器',3:u'成员服务器',4:'BDC',5:'PDC',6:'Unknown'}[i.DomainRole],i.domain))
		# 硬盘
		disk_info = []
		for i in w.win32_diskdrive():
			if i.InterfaceType == 'USB':
				continue
			disk_info.append('%s (%dG)' %(i.caption, int(i.size) / 1000000000))
		self.disk.AppendText('\n'.join(disk_info))
		#cpu
		cpu_info = []
		for i in w.win32_processor():
			cpu_info.append(i.Name)
		self.cpu.AppendText('\n'.join(cpu_info))
		# self.cpu.AppendText(cpuidpy.model_name.strip()+' %sMHz' %i.CurrentClockSpeed)

		# for i in w.win32_cachememory():
		# 	self.cpu.AppendText(' %s:%d'%(i.purpose,i.maxcachesize))
		# 	#print i.SocketDesignation, '  L2cache%sKb' %i.l2cachesize
		# 	#print i.DeviceID
		for i in w.win32_baseboard():
			self.board.AppendText('%s %s' % (i.Product, i.Version)) #i.Manufacturer

		ip_info = []
		mac_info = []
		for i in w.Win32_NetworkAdapterConfiguration():
			if i.IPEnabled == True:
				ip_info.extend(i.IPAddress)
				mac_info.append(i.MACAddress)
		self.ip.AppendText('\n'.join(ip_info))
		self.mac.AppendText('\n'.join(mac_info))

		display_info = []
		for i in w.Win32_VideoController():
			display_info.append(i.Name)
		self.display.AppendText('\n'.join(display_info))

		# 显示器
		monitor_info = []
		for i in w.Win32_DesktopMonitor():
			monitor_info.append(i.Name)
		self.monitor.AppendText('\n'.join(monitor_info))

	def on_btn(self, e):
		self.model = e.GetEventObject().GetLabelText()

	def toexl(self,evt):
		if self.name.GetValue().strip()=='':
			dlg=wx.MessageDialog(self, '姓名不能为空！', 'PC数据收集测试版',wx.OK | wx.ICON_INFORMATION)
			dlg.ShowModal()
			dlg.Destroy()
			self.name.SetFocus()
			return

		exl=win32com.client.Dispatch('Excel.Application')
		book = exl.Workbooks.Open(path.join(path.dirname(path.abspath(__file__)), 'out.xlsx'))
		sht = book.Worksheets(1)
		nrows=sht.usedrange.rows.count
		sht.Cells(nrows+1, 1).value = self.dept.GetValue()
		sht.Cells(nrows+1, 2).value = self.name.GetValue()
		sht.Cells(nrows+1, 3).value=self.board.GetValue()
		sht.Cells(nrows+1, 4).value=self.cpu.GetValue()
		sht.Cells(nrows+1, 5).value=self.display.GetValue()
		sht.Cells(nrows+1, 6).value=self.disk.GetValue()
		sht.Cells(nrows+1, 7).value=self.mem_size.GetValue()
		sht.Cells(nrows+1, 8).value=self.ip.GetValue()
		sht.Cells(nrows+1, 9).value=self.mac.GetValue()
		sht.Cells(nrows+1, 10).value=self.monitor.GetValue()
		sht.Cells(nrows+1, 11).value=self.model
		sht.Cells(nrows+1, 12).value=self.other.GetValue()
		#book.Save()
		book.Close(SaveChanges=1)
		dlg=wx.MessageDialog(self, self.name.GetValue() + '的电脑配置信息保存成功！', 'PC数据收集测试版', wx.OK | wx.ICON_INFORMATION)
		dlg.ShowModal()
		dlg.Destroy()

if __name__=='__main__':
	app=wx.App()
	main_frame().Show()
	app.MainLoop()






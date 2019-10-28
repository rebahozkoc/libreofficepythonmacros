#!/usr/bin/python
# -*- coding: utf-8 -*-
#
#	loook.py 0.6.8 (2010-12-04) - Search in OpenOffice.org files
#
#	Changes 0.6.7 -> 0.6.8: search embedded files, thanks to Hanna Ruotsalainen
#	Changes 0.6.6 -> 0.6.7: fixes to make loook work with Python 3.1.1 when started with parameters
#	Changes 0.6.5 -> 0.6.6: made loook work with Python 3.1.1
#
#	Homepage: http://www.danielnaber.de/loook/
#
#	Copyright (C) 2003-2010 Daniel Naber, Lutz Haseloff
#
#	This program is free software; you can redistribute it and/or
#	modify it under the terms of the GNU General Public License
#	as published by the Free Software Foundation; either version 2
#	of the License, or (at your option) any later version.
#
#	This program is distributed in the hope that it will be useful,
#	but WITHOUT ANY WARRANTY; without even the implied warranty of
#	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#	GNU General Public License for more details.
#
#	You should have received a copy of the GNU General Public License
#	along with this program; if not, write to the Free Software
#	Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

# TODO:
#	-use better defaults for the file/dir dialogs:
#		askdirectory(initialdir="..."...
#	-support KOffice and Abiword?
#	-show title (use HListbox?)

import configparser
import codecs
import os
import re
import string
import time
import zipfile
import locale

from tkinter import * 
import tkinter.filedialog
import tkinter.messagebox

class Application:

	CONFIGFILE = ".loook.cfg"
	
	def __init__(self, master=None):
		"""Load configuration or use sensible default values."""
		self.master = master
		self.stopped = 0
		self.ooo_path_str = None
		self.search_path_str = None
		config_path = None
		if os.getenv('USERPROFILE'):
			config_path = os.getenv('USERPROFILE')
		elif os.getenv('HOME'):
			config_path = os.getenv('HOME')
		elif os.name == 'dos':
			config_path = "c:/"
		if config_path:
			self.configfile = os.path.join(config_path, self.CONFIGFILE)
			self.config = configparser.ConfigParser()
			self.config.read(self.configfile)
			try:
				self.ooo_path_str = self.config.get("General", "ooo_path")
				self.search_path_str = self.config.get("General", "search_path")
			except configparser.NoSectionError:
				pass
		else:
			print >> sys.stderr, "Cannot find home directory, settings will not be saved."
			self.configfile = None
		self.createWidgets()
		return

	def createWidgets(self):
		"""Build and show the GUI elements."""
		if os.name != 'nt':
			Label(self.master, text="Viewer:").grid(row=0, sticky=E)
		Label(self.master, text="Search path:").grid(row=1, sticky=E)
		Label(self.master, text="Search terms:").grid(row=2, sticky=E)
		Label(self.master, text="Mode:").grid(row=3, sticky=E)
		Label(self.master, text="Matches:").grid(row=4, sticky=N+E)

		self.ooo_path = Entry(self.master)
		if self.ooo_path_str:
			self.ooo_path.insert(END, self.ooo_path_str)
		else:
			self.ooo_path.insert(END, "soffice")
		self.ooo_path_b = Button(self.master)
		self.ooo_path_b.bind("<Button-1>", self.selectOOoPath)
		self.ooo_path_b["text"] = ">"

		self.search_path = Entry(self.master)
		if len(sys.argv) >= 2:
			self.search_path.insert(END, sys.argv[1])
		elif self.search_path_str:
			self.search_path.insert(END, self.search_path_str)
		else:
			self.search_path.insert(END, os.getcwd())
		self.search_path_b = Button(self.master)
		self.search_path_b.bind("<Button-1>", self.selectSearchPath)
		self.search_path_b["text"] = ">"

		self.search_query = Entry(self.master)
		self.search_query.bind('<Return>', self.startSearch)
		if len(sys.argv) >= 2:
			lang, enc = locale.getdefaultlocale()
			self.search_query.insert(END, ' '.join(sys.argv[2:]))
		self.search_query.focus()

		self.mode_button = Button(self.master)
		self.mode_button.bind("<Button-1>", self.popupMode)
		self.mode_button["text"] = "AND"
		self.mode_menu = Menu(self.master, tearoff=0)
		self.mode_menu.add_command(label="AND", command=self.setModeAND)
		self.mode_menu.add_command(label="OR", command=self.setModeOR)
		self.mode_menu.add_command(label="Phrase", command=self.setModePhrase)

		pad = 1
		if os.name != 'nt':
			self.ooo_path.grid(columnspan=2, row=0, column=1, sticky=E+W, pady=pad, padx=pad)
			self.ooo_path_b.grid(row=0, column=3, sticky=E+W, pady=pad, padx=pad)
		self.search_path.grid(columnspan=2, row=1, column=1, sticky=E+W, pady=pad, padx=pad)
		self.search_path_b.grid(row=1, column=3, sticky=E+W, pady=pad, padx=pad)
		self.search_query.grid(columnspan=3, row=2, column=1, sticky=E+W, pady=pad, padx=pad)
		self.mode_button.grid(columnspan=3, row=3, column=1, sticky=W, pady=pad, padx=pad)

		self.scrollbar = Scrollbar(self.master)
		self.scrollbar.grid(row=4, column=3, sticky=N+S, pady=pad, padx=pad)
		self.listbox = Listbox(self.master, yscrollcommand=self.scrollbar.set)
		self.listbox.bind('<Double-Button-1>', self.showDoc)
		self.listbox.grid(columnspan=2, row=4, column=1, sticky=E+W+S+N, pady=pad, padx=pad)
		self.scrollbar.config(command=self.listbox.yview)		
		
		f = Frame(self.master)
		self.search = Button(f)
		self.search["text"] = "Search"
		self.search["command"] = self.startSearch
		self.search.pack(side=LEFT)
		self.quit_button = Button(f)
		self.quit_button["text"] = "Quit"
		self.quit_button["command"] = self.doQuit
		self.quit_button.pack(side=RIGHT)
		self.stop_button = Button(f)
		self.stop_button["text"] = "Stop"
		self.stop_button["command"] = self.stop
		self.stop_button["state"] = DISABLED
		self.stop_button.pack(side=RIGHT)
		f.grid(row=5, columnspan=2, column=2, sticky=E, pady=pad, padx=pad)

		self.status = Label(self.master, text="", bd=1, relief=SUNKEN, anchor=W)
		self.status.config(text="Ready.")
		self.status.grid(row=6, columnspan=4, column=0, sticky=E+W, pady=pad, padx=pad)
		return

	def doQuit(self):
		"""Save configuration, the quit."""
		self.saveConfig()
		self.master.quit()
		return

	def saveConfig(self):
		"""Save path settings in configuration file in the user's HOME."""
		if self.configfile:
			file = codecs.open(self.configfile, "w", "utf-8")
			if not self.config.has_section("General"):
				self.config.add_section("General")
			file.write("[General]\n")
			file.write("ooo_path=%s\n" % self.ooo_path.get())
			file.write("search_path=%s\n" % self.search_path.get())
			file.close()
		return

	def selectOOoPath(self, event):
		d = tkinter.filedialog.askopenfilename()
		self.ooo_path.delete(0, END)
		self.ooo_path.insert(END, d)
		return

	def selectSearchPath(self, event):
		d = tkinter.filedialog.askdirectory()
		self.search_path.delete(0, END)
		self.search_path.insert(END, d)
		return

	def setMode(self, mode):
		self.mode = mode
		self.mode_button["text"] = mode
		return

	def setModeAND(self):
		self.setMode("AND")
		return

	def setModeOR(self):
		self.setMode("OR")
		return

	def setModePhrase(self):
		self.setMode("Phrase")
		return

	def popupMode(self, event):
		try:
			self.mode_menu.tk_popup(event.x_root, event.y_root, 0)
		finally:
			# make sure to release the grab (Tk 8.0a1 only)
			self.mode_menu.grab_release()
		return
		
	def stop(self):
		self.stopped = 1
		return
		
	def showDoc(self, event):
		"""Start OOo to view the file. This method lacks 
		error handling (TODO)."""
		items = event.widget.curselection()
		#try: items = map(string.atoi, items)
		#except ValueError: pass
		filename = "%s%s" % (self.search_path.get(), event.widget.get(items[0]))
		filename = os.path.normpath(filename)
		if os.name == 'nt':
			try:
				os.startfile(filename)
			except:
				print("Warning: File could not be opened. - %s" % filename)
		else:
			prg = self.ooo_path.get()
			if not prg and os.name != 'nt':
				tkinter.messagebox.showwarning('Error', 'Set viewer first.')
			else:
				#fixme: feedback if this fails
				#fixme: feedback while OOo starts?
				#fixme: fails if there's a quote in the prg variable (does this matter?)
				#fixme: does a filename with quotes work on windows?
				filename = filename.replace('"', '\\"')
				if os.name == 'dos':
					filename = filename.replace('/', '\\')
					prg = prg.replace('/', '\\')
				cmd = "\"%s\" \"%s\" &" % (prg, filename)
				if os.name == 'dos':
					cmd = "\"%s\" \"%s\"" % (prg, filename)
				self.status.config(text="Starting viewer...")
				print(cmd)
				try:
					res = os.system(cmd)
				except UnicodeError:
					res = os.system(cmd)
				if res != 0:
					# don't show a dialog, this check might not be system-indepenent:
					print("Warning: Command returned code != 0: %s" % cmd)
		return
		
	def removeXMLMarkup(self, s, replace_with_space):
		s = re.compile("<!--.*?-->", re.DOTALL).sub('', s)
		repl = ''
		if replace_with_space:
			repl = ' '
		s = re.compile("<[^>]*>", re.DOTALL).sub(repl, s)
		return s

	def match(self, query, docstring):
		mode = self.mode_button["text"]
		if mode == "Phrase":
			# match only documents that contain the phrase:
			regex = re.compile(re.escape(query.lower()), re.DOTALL)
			if regex.search(docstring):
				return 1
		else:
			parts = re.split("\s+", query.strip())
			if mode == "AND":
				# match only documents that contain all words:
				match = 1
				for part in parts:
					regex = re.compile(re.escape(part.lower()), re.DOTALL)
					if not regex.search(docstring):
						match = 0
						break
				return match
			elif mode == "OR":
				# match documents that contain at leats one word:
				match = 0
				for part in parts:
					regex = re.compile(re.escape(part.lower()), re.DOTALL)
					if regex.search(docstring):
						match = 1
						break
				return match
			else:
				print("Error: unknown search mode '%s'" % mode)
		return 0

	def processFile(self, filename, query):
		suffix = self.getSuffix(filename)
		try:
			# Handle OpenOffice.org files:
			if suffix in ('sxw', 'stw',				# OOo 1.x swriter
					'sxc', 'stc',					# OOo 1.x scalc
					'sxi', 'sti'					# OOo 1.x simpress
					'sxg',							# OOo 1.x master document
					'sxm',							# OOo 1.x formula
					'sxd', 'std',					# OOo 1.x sdraw
					'odt', 'ott',					# OOo 2.x swriter
					'odp', 'otp',					# OOo 2.x simpress
					'odf',							# OOo 2.x formula
					'odg', 'otg',					# OOo 2.x sdraw
					'ods', 'ots'					# OOo 2.x scalc
					):
				zip = zipfile.ZipFile(filename, "r")
				content = ""
				docinfo = ""
				try:
					# TODO: are all OOo files utf-8?
					# search embedded files:
					filelist = zip.namelist()
					for filename in filelist:
						if filename.endswith("content.xml"):
							content += str(zip.read(filename), 'utf-8')
					# TODO: is replace_with_space=0 correct?
					content = self.removeXMLMarkup(content, replace_with_space=0)
					docinfo = str(zip.read("meta.xml"), 'utf-8')
					docinfo = self.removeXMLMarkup(docinfo, replace_with_space=0)
					self.ooo_count = self.ooo_count + 1
				except KeyError as err:
					print("Warning: %s not found in '%s'" % (err, filename))
					return None
				title = ""
				title_match = re.compile("<dc:title>(.*?)</dc:title>", re.DOTALL|re.IGNORECASE).search(docinfo)
				if title_match:
					title = title_match.group(1)
				if self.match(query, "%s %s" % (content.lower(), docinfo.lower())):
					return (filename, title)
			# Handle KWord files:
			# TODO: this will need its own viewer
			#elif suffix == 'kwd':
			#	zip = zipfile.ZipFile(filename, "r")
			#	content = unicode(zip.read("maindoc.xml"), 'utf-8')
			#	content = self.removeXMLMarkup(content, replace_with_space=0)
			#	docinfo = unicode(zip.read("documentinfo.xml"), 'utf-8')
			#	docinfo = self.removeXMLMarkup(docinfo, replace_with_space=0)
			#	title = ""
			#	title_match = re.compile("<about.*?<title>(.*?)</title>", re.DOTALL|re.IGNORECASE).search(docinfo)
			#	if title_match:
			#		title = title_match.group(1)
			#	if regex.search(content.lower()) or regex.search(docinfo.lower()):
			#		return (filename, title)
		except zipfile.BadZipfile as err:
			print("Warning: Supposed ZIP file '%s' could not be opened: %s" % (filename, err))
			#if suffix == 'kwd':
			#	print("Note that the old KOffice (< 1.2) file format is not supported.")
		except IOError as err:
			print("Warning: File '%s' could not be opened: %s" % (filename, err))
		return None

	def startSearch(self, event=None):
		self.stopped = 0
		self.last_update = 0
		self.match_count = 0
		self.ooo_count = 0
		self.listbox.delete(0, END)
		self.stop_button["state"] = NORMAL
		self.quit_button["state"] = DISABLED
		if not os.path.exists(self.search_path.get()):
			tkinter.messagebox.showwarning('Error', 
				'Path "%s" doesn\'t exist' % self.search_path.get())
		else:
			start_time = time.time()
			self.recursiveSearch(self.search_path.get())
			duration = time.time() - start_time
			#print("time=%.2f" % duration)
			count_str = "in %d files" % self.ooo_count
			if self.stopped:
				self.status.config(text="%d matches so far %s (search stopped)" % (self.match_count, count_str))
				self.stopped = 0
			else:
				self.status.config(text="%d matches %s" % (self.match_count, count_str))
		self.stop_button["state"] = DISABLED
		self.quit_button["state"] = NORMAL
		return

	def getSuffix(self, filename):
		suffix_match = re.compile(".*\.(.*)").match(filename)
		if suffix_match:
			suffix = suffix_match.group(1).lower()
		else:
			suffix = None
		return suffix

	def recursiveSearch(self, directory):
		len_limit = 15		# avoid resizing window
		dir_part = os.path.split(directory)[1]
		if len(dir_part) > len_limit:
			dir_part = "%s..." % dir_part[0:len_limit]
		#print("'%s'" % dir_part)
		self.status.config(text="Searching in %s" % dir_part)
		try:
			dir_content = os.listdir(directory)
			dir_content.sort(key=str.lower) 
		except OSError as err:
			print("Warning: %s: %s" % (directory, err))
			return
		except UnicodeError as err:
			print("Warning: Unicode problem with directory name...")
			return
		for filename in dir_content:
			if self.stopped:
				return
			filename = os.path.join(directory, filename)
			if os.path.isfile(filename):
				match = self.processFile(filename, self.search_query.get())
				update_interval = 0.05
				time_diff = time.time() - self.last_update
				if time_diff > update_interval:
					self.master.update()
					self.last_update = time.time()
				if match:
					title = match[1]
					if not title:
						title = "Untitled"
					# TODO: reactive title:
					#self.listbox.insert('end', "%s - %s" % (filename, title))
					display_filename = filename.replace(self.search_path.get(), '')
					self.listbox.insert('end', "%s" % display_filename)
					self.match_count = self.match_count + 1
			elif os.path.isdir(filename) and not os.path.islink(filename):
				self.recursiveSearch(filename)
		return

if __name__ == "__main__":
	if len(sys.argv) >= 2 and (sys.argv[1] == '--help' or sys.argv[1] == '-h'):
		print("Usage: loook.py [-h|--help] [directory] [search term]...")
		sys.exit(1)
	root = Tk()
	root.minsize(380, 200)
	version = "0.6.5"
	root.title("loook.py %s - OpenOffice.org File Finder" % version)
	root.columnconfigure(1, weight=1)
	root.rowconfigure(4, weight=1)
	root.columnconfigure(1, weight=1)
	root.rowconfigure(5, weight=0)

	app = Application(root)
	root.protocol("WM_DELETE_WINDOW", app.doQuit)
	root.mainloop()

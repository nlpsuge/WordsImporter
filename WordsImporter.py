# -*- coding: utf-8 -*-
#
# Version: 1.0
import os

import httplib
import thread

from PyQt4.QtCore import SIGNAL
from PyQt4 import QtGui
import anki, aqt, re

import sys

sys.path.append('/usr/lib/python2.7/site-packages')
sys.path.append('/usr/lib64/python2.7')

# need to run `pip install selenium`
from selenium import webdriver

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import subprocess
import traceback

# pip install python-magic
import magic

# pip install selenium
import selenium


class ImportWords:

    # if not a word
    # or if not a word that not contains any whitespace or '-'
    # https://docs.python.org/2/library/re.html
    wordRegex = re.compile(r'[^A-Za-z\s-]+?')

    # constant start
    shortCut = 'Ctrl+W'

    FIREFOX_BINARY = '/usr/bin/firefox'
    EXECUTABLE_GECKODRIVER_PATH = 'the_path_of_geckodriver'

    generateHtml5 = True

    # youdao
    youdaoBaseurl = "http://www.youdao.com/"
    youdaoSearchurl = youdaoBaseurl + "w/"
    youdaoMediaURL = youdaoBaseurl + "example/mdia/%s/#keyfrom=dict.main.moremedia"
    #   %d = 1:UK/British, 2:US/American
    youdaoVoiceURL = "http://dict.youdao.com/dictvoice?type=%d&audio=%s"

    # fileds
    youdaoXpaths = {
        'phoneticSymbolUK': ".//*[@id='phrsListTab']/h2/div/span[1]/span",
        'phoneticSymbolUS': ".//*[@id='phrsListTab']/h2/div/span[2]/span",
        'simpleMeaning': ".//*[@id='phrsListTab']/div[@class='trans-container']/ul",
        '21CenturyDictionaryRoot': ".//*[@id='authDictTrans']",
        '21CenturyDictionaryShowMore': ".//*[@id='authDictTrans']/div[1]",
        'phraseBar': ".//*[@id='eTransform']/h3/span/a[@rel='#wordGroup']/span",
        'phrase': ".//*[@id='wordGroup']",
        'phraseShowMore': ".//*[@id='wordGroup']/div",
        'nativeSpeakerExamplesBar': "//div[@id='examples']/h3/span/a[2]/span",
        'nativeSpeakerExamples': ".//*[@id='originalSound']/ul/li[%d]",
        'nativeSpeakerVideo': ".//*[@id='originalSound']/ul/li[%d]/div/a[1]",
        'synonymBar': ".//*[@id='eTransform']/h3/span/a[@rel='#synonyms']",
        'synonym': ".//*[@id='synonyms']/ul",
        'synonymMore': "",
        'differencesBar': ".//*[@id='eTransform']/h3/span/a[@rel='#discriminate']/span",
        'differences': ".//*[@id='discriminate']",
    }

    # 91dict TODO
    dict91Baseurl = 'http://www.91dict.com/words'
    dict91Searchurl = dict91Baseurl + "?w/="
    dict91Xpath = {'exampleSentenceBar': "//nav/div/div/ul/li[@class='lastlili current']",

                   }

    # constant end

    def __init__(self):
        def setup_Menu(_browser):
            self.ankiCollectionMediaPath = os.path.join(_browser.mw.pm.profileFolder(), "collection.media")
            menu = QtGui.QMenu('WordsImporter', _browser.form.menubar)
            _browser.form.menubar.addMenu(menu)

            def append_Munu(_text, how):
                action = QtGui.QAction(_text, menu)
                # set shot cut
                from PyQt4.QtGui import QKeySequence
                action.setShortcut(QKeySequence(self.shortCut))
                _browser.connect(action, SIGNAL('triggered()'), lambda s=_browser: self._run(s, how))
                menu.addAction(action)

            append_Munu('Import selected word(s)...', 'useless parameter')

        # def setup_context_menu(self):
        #
        #     def insert_search_menu_action(anki_web_view, m):
        #
        #         print('awv of setup_context_menu ->' + str(anki_web_view))
        #
        #         word = anki_web_view.page().selectedText()
        #
        #         action = m.addAction('Retrieve "%s" word...' % word)
        #
        #         action.connect(action, SIGNAL("triggered()"),
        #                   lambda wv = anki_web_view: do_search_for_selection(wv))
        #
        #     return insert_search_menu_action

        anki.hooks.addHook('browser.setupMenus', setup_Menu, )
        # anki.hooks.addHook("AnkiWebView.contextMenuEvent", setup_context_menu())
        # anki.hooks.addHook('EditorWebView.contextMenuEvent', setup_context_menu())

    def _run(self, _browser, how):
        _notes = [_browser.mw.col.getNote(note_id) for note_id in _browser.selectedNotes()]
        if not _notes:
            aqt.utils.showWarning('Please select one/some item(s).', _browser)
            return
        _browser.model.beginReset()
        self._done(_browser, _notes, how)
        _browser.model.endReset()

    def _done(self, _browser, _notes, how):
        count_i = 0
        total_i = len(_notes)
        count_f = float(0)
        total_f = float(len(_notes))
        for note in _notes:
            # trim leading or tail whitespace
            word = unicode.strip(note['Front'])
            if self.wordRegex.search(word):
                continue

            try:
                self.importWords(note, word)

            except Exception as e:
                print('Exception occurs when execute _done()')
                print(traceback.format_exc())
            finally:
                count_f += 1
                count_i += 1
                print('Progress: ' + str((count_f / total_f) * 100) + '%'
                      + ' (' + str(count_i) + '/' + str(total_i) + ')')

    # selenium start
    @staticmethod
    def element_exist(xpath):
        try:
            mydriver.find_element_by_xpath(xpath)
            return True
        except:
            return False

    def runFireFox(self):
        if 'mydriver' not in globals():
            print('opening browser')
            global mydriver
            fp = webdriver.FirefoxProfile()
            # forbid to display the images
            fp.set_preference("permissions.default.image", 2)
            # block all cookies to resolve some elements can not be clicked
            # https://developer.mozilla.org/en-US/docs/Mozilla/Cookies_Preferences
            fp.set_preference("network.cookie.cookieBehavior", 2)
            # unset proxy
            # type -> Direct = 0, Manual = 1, PAC = 2, AUTODETECT = 4, SYSTEM = 5
            fp.set_preference("network.proxy.type", 0)
            options = webdriver.FirefoxOptions()
            # set headless
            # options.headless = True
            mydriver = webdriver.Firefox(executable_path=self.EXECUTABLE_GECKODRIVER_PATH,
                                         firefox_binary=self.FIREFOX_BINARY,
                                         firefox_profile=fp,
                                         options=options)
            mydriver.minimize_window()

    # def checkStateOfFirefox(self):
    #     # https://stackoverflow.com/questions/18619121/in-java-best-way-to-check-if-selenium-webdriver-has-quit
    #     return False

    def doSearch(self, word):
        searchword = self.youdaoSearchurl + word
        print('Searching ' + searchword)
        mydriver.get(searchword)

    def getIdentifiedId(self):
        return 'youdao'

    def click_element(self, xpath):
        from selenium.webdriver.support.wait import WebDriverWait
        from selenium.webdriver.support import expected_conditions
        from selenium.webdriver.common.by import By
        wait = WebDriverWait(mydriver, 5)

        wait.until(expected_conditions.presence_of_element_located((By.XPATH, xpath)))

        # move to element to fix 'Element is not clickable at point'
        # see: https://stackoverflow.com/questions/11908249/debugging-element-is-not-clickable-at-point-error
        mydriver.execute_script('arguments[0].click();', mydriver.find_element_by_xpath(xpath))

        wait.until(expected_conditions.element_to_be_clickable((By.XPATH, xpath))).click()

    # youdao voice
    def downloadYaodaoVoice(self, word):
        # 1:UK/British, 2:US/American
        voiceOfUKUrl = self.youdaoVoiceURL % (1, word)
        fileNameUK = self.getIdentifiedId() + '_' + word.replace(' ', '-') + '_UK.mp3'
        fileUK = self.buildFilePath(fileNameUK)

        self.download(voiceOfUKUrl, fileUK, checkFileExists=True)

        voiceOfUSUrl = self.youdaoVoiceURL % (2, word)
        fileNameUS = self.getIdentifiedId() + '_' + word.replace(' ', '-') + '_US.mp3'
        fileUS = self.buildFilePath(fileNameUS)

        self.download(voiceOfUSUrl, fileUS, checkFileExists=True)

    def downloadYaodaoVoiceWithinThread(self, word):
        thread.start_new_thread(self.downloadYaodaoVoice, (word,))

    # TODO
    def buildYaodaoVoiceTag(self, word):
        self.downloadYaodaoVoiceWithinThread(word)

        fileNameUK = self.getIdentifiedId() + '_' + word.replace(' ', '-') + '_UK.mp3'
        fileNameUS = self.getIdentifiedId() + '_' + word.replace(' ', '-') + '_US.mp3'
        # voice_tag = ''
        # if Path(buildFilePath(fileNameUK)).exists():
        voice_tag = '英 [sound:%s] ' % fileNameUK
        # if Path(buildFilePath(fileNameUS)).exists():
        voice_tag = voice_tag + '美 [sound:%s]' % fileNameUS
        return voice_tag.strip()

    # phonetic symbol
    def getPhoneticSymbol(self):
        ps = ''
        if self.element_exist(self.youdaoXpaths['phoneticSymbolUK']):
            ps_uk = '英 ' + mydriver.find_element_by_xpath(self.youdaoXpaths['phoneticSymbolUK']).text
            ps = ps + ps_uk + ' '
        if self.element_exist(self.youdaoXpaths['phoneticSymbolUS']):
            ps_us = '美 ' + mydriver.find_element_by_xpath(self.youdaoXpaths['phoneticSymbolUS']).text
            ps = ps + ps_us
        ps = ps.strip()
        return ps

    def getSimpleMeaning(self):
        if self.element_exist(self.youdaoXpaths['simpleMeaning']):
            sm = mydriver.find_element_by_xpath(self.youdaoXpaths['simpleMeaning']).text
            return sm

    # 21 century dictionary
    def get21CenturyDictionary(self):
        if self.element_exist(self.youdaoXpaths['21CenturyDictionaryRoot']):
            if self.element_exist(self.youdaoXpaths['21CenturyDictionaryShowMore']):
                self.click_element(self.youdaoXpaths['21CenturyDictionaryShowMore'])

            century_dictionary_root = '21CenturyDictionaryRoot'
            element_removal_xpath = './/div[@class="more"]'

            return self.get_html_from_pagesource(century_dictionary_root, element_removal_xpath)

    def get_html_from_pagesource(self, century_dictionary_root, element_removal_xpath=None):
        html = mydriver.find_element_by_xpath(self.youdaoXpaths[century_dictionary_root]) \
            .get_attribute('innerHTML')

        if element_removal_xpath:
            from lxml import etree
            passer = etree.HTMLParser(encoding='utf-8', remove_comments=True, remove_blank_text=True)
            import StringIO
            tree = etree.parse(StringIO.StringIO(html), passer)

            e = tree.xpath(element_removal_xpath)
            if len(e) == 1:
                e[0].getparent().remove(e[0])

            return etree.tostring(tree.getroot(), encoding='utf-8', pretty_print=True, method="html")
        else:
            return html

    # phrase
    def get_phrase_if_has_phrase_bar(self):

        if self.element_is_visible(self.youdaoXpaths['phrase']):
            return self.get_phrase()

        if self.element_exist(self.youdaoXpaths['phraseBar']):
            self.click_element(self.youdaoXpaths['phraseBar'])

            return self.get_phrase()

    def get_phrase(self):
        if self.element_exist(self.youdaoXpaths['phraseShowMore']):
            self.click_element(self.youdaoXpaths['phraseShowMore'])

        phrase = unicode.replace(
            mydriver.find_element_by_xpath(self.youdaoXpaths['phrase']).text,
            '收起词组短语',
            '')
        return phrase

    # synonym
    def getSynonym(self):
        if self.element_is_visible(self.youdaoXpaths['synonym']):
            synonym = mydriver.find_element_by_xpath(self.youdaoXpaths['synonym']).text
            return synonym

        if self.element_exist(self.youdaoXpaths['synonymBar']):
            self.click_element(self.youdaoXpaths['synonymBar'])
            synonym = mydriver.find_element_by_xpath(self.youdaoXpaths['synonym']).text
            return synonym

    # differences TODO
    def getDifferences(self):
        if self.element_is_visible(self.youdaoXpaths['differences']):
            return mydriver.find_element_by_xpath(self.youdaoXpaths['differences']).text

        if self.element_exist(self.youdaoXpaths['differencesBar']):
            self.click_element(self.youdaoXpaths['differencesBar'])
            return self.get_html_from_pagesource('differences')

    def downloadAndConvertFlv2Mp4(self, href, filePath, checkFileExists=False):
        try:
            self.download(href, filePath, checkFileExists)

            fileType = magic.from_file(filePath, mime=True)

            if fileType == 'video/x-flv':
                newFileName = unicode.replace(filePath, '.flv', '.mp4')

                if self.fileNotExists(newFileName):
                    # https://superuser.com/questions/483597/converting-flv-to-mp4-using-ffmpeg-and-preserving-the-quality
                    # https://stackoverflow.com/questions/15264508/convert-the-videos-of-any-formatflv-3gp-mxf-etc-to-mp4-in-django-using-pytho
                    shellCommand = ('ffmpeg -i "%s" -c:v libx264 -crf 23 -c:a aac -q:a 100 -strict -2 "%s" -y') \
                                   % (filePath, newFileName)
                    print('shellCommand -> ' + shellCommand)
                    subprocess.call(shellCommand, shell=True)

                if self.fileNotExists(filePath) is False:
                    print('Deleting ' + filePath)
                    # delete origin file
                    import os
                    os.remove(filePath)
        except:
            print(traceback.format_exc())

    def downloadAndConvertFlv2Mp4WithinThread(self, href, filePath, checkFileExists):
        thread.start_new_thread(self.downloadAndConvertFlv2Mp4, (href, filePath, checkFileExists,))

    def getHtml5Video(self, href, filePath, fileName, checkFileExists=False):
        self.downloadAndConvertFlv2Mp4WithinThread(href, filePath, checkFileExists)
        videoTag = '\n<div><video width="400" controls="controls">\n' \
                   '<source src="%s" type="video/mp4">\n' \
                   '</source></video></div>' % unicode.replace(fileName, '.flv', '.mp4')
        return videoTag

    # Return True if file 'absolutePath' not exists, otherwise return False
    # If the file 'absolutePath' can not to be use please delete it at first, after that try again
    def fileNotExists(self, absolutePath):

        from pathlib import Path
        path = Path(absolutePath)
        if path.exists():
            print('file [%s] is exists' % absolutePath)
            return False
        return True

    # sentence && videos
    def getSentencesAndVideos(self, word):
        sv = ''
        mydriver.get(self.youdaoMediaURL % word)
        for i in range(1, 7):
            if self.element_exist(self.youdaoXpaths['nativeSpeakerExamples'] % i):
                nsvs = mydriver.find_element_by_xpath(self.youdaoXpaths['nativeSpeakerExamples'] % i)

                sv = sv + nsvs.text
                if self.element_exist(self.youdaoXpaths['nativeSpeakerVideo'] % i):
                    nsv = mydriver.find_element_by_xpath(self.youdaoXpaths['nativeSpeakerVideo'] % i)
                    href = unicode.replace(
                        nsv.get_attribute('href'),
                        'http://www.youdao.com/simplayer.swf?movie=', '')
                    fileName = self.getIdentifiedId() + '_' + word.replace(' ', '-') + "_%d.flv" % i
                    ankiCollectionMediaFilePath = self.buildFilePath(fileName)

                    if self.generateHtml5:
                        newFileName = unicode.replace(fileName, '.flv', '.mp4')
                        sv = sv + '[sound:' + newFileName + ']\n'
                    else:
                        sv = sv + '[sound:' + fileName + ']\n'

                    if self.generateHtml5:
                        sv = sv + self.getHtml5Video(href, ankiCollectionMediaFilePath, fileName, checkFileExists=True)
                    else:
                        self.downloadWithinThread(href, ankiCollectionMediaFilePath, checkFileExists=True)

                sv = sv + '\n'
        return sv

    def buildFilePath(self, fileName):
        ankiCollectionMediaFilePath = self.ankiCollectionMediaPath + '/' + fileName
        return ankiCollectionMediaFilePath

    # 'checkFileExists' declares if check existence of 'filePath'
    def download(self, href, filePath, checkFileExists=False):
        if checkFileExists is False or (checkFileExists is True and self.fileNotExists(filePath) is True):
            try:
                print('downloading %s to %s' % (href, filePath))
                import requests
                response = requests.get(href)
                out_file = open(filePath, 'wb')
                out_file.write(response.content)
                out_file.flush()
            except httplib.IncompleteRead as ir:
                print(traceback.format_exc())
                print('Try again...')
                import os
                # File must be damaged, so delete it at first if checkFileExists is True
                if checkFileExists is True:
                    os.remove(filePath)
                self.download(href, filePath, checkFileExists)
            except:
                print(traceback.format_exc())
            finally:
                out_file.close()

    def downloadWithinThread(self, href, filePath, checkFileExists=False):
        thread.start_new_thread(self.download, (href, filePath, checkFileExists))

    # selenium end

    def importWords(self, note, word):
        try:
            self.runFireFox()
            self.doSearch(word)

            # replace all '\n' to '<br/>', cos '\n' will be trimed after inserted into SQLite
            # set ps
            note['phonetic symbol'] = self.replace_newline2br(str(self.getPhoneticSymbol()))
            print('phonetic symbol -> ' + note['phonetic symbol'])
            # set youdao voice
            voice_tag = str(self.buildYaodaoVoiceTag(word))
            if voice_tag != '':
                note['voice'] = voice_tag
            print('voice -> ' + note['voice'])
            # set simple meaning
            note['simple meaning'] = self.replace_newline2br(str(self.getSimpleMeaning()))
            print('simple meaning -> ' + note['simple meaning'])
            # set 21
            note['21 century dictionary'] = str(self.get21CenturyDictionary())
            print('21 century dictionary -> ' + note['21 century dictionary'])
            # set phrase
            note['phrase'] = self.replace_newline2br(str(self.get_phrase_if_has_phrase_bar()))
            print('phrase -> ' + note['phrase'])

            # set synonym
            note['synonym'] = self.replace_newline2br(str(self.getSynonym()))
            print('synonym -> ' + note['synonym'])
            # set differences
            note['differences'] = str(self.getDifferences())
            print('differences -> ' + note['differences'])

            # set sentence and get video, put it at last
            note['sentence'] = self.replace_newline2br(str(self.getSentencesAndVideos(word)))
            print('sentence -> ' + note['sentence'])

        except httplib.BadStatusLine:
            print('httplib.BadStatusLine occurs when open ' + word)
            print(traceback.format_exc())
            # re-do if an exception of httplib.BadStatusLine occur
            self.importWords(note, word)
        except selenium.common.exceptions.SessionNotCreatedException as e:
            print('opening url in the new tab')
            print(traceback.format_exc())

            # reset global variable mydriver
            globals().__delitem__('mydriver')
            self.importWords(note, word)

        except selenium.common.exceptions.NoSuchWindowException as e:
            print('reset global variavle mydriver, we will open a new browser instance')
            print(traceback.format_exc())

            # try to switch a new tab then open url in this tab rather than open a new browser instance
            hs = mydriver.window_handles
            mydriver.switch_to.window(hs[0])
            self.importWords(note, word)
        except selenium.common.exceptions.WebDriverException as wde:
            print(traceback.format_exc())
            self.importWords(note, word)
        except Exception as e:
            print(traceback.format_exc())
        finally:
            note.flush()

    def doWithinThread(self, word):
        # does not work within a thread as expected...?
        thread.start_new_thread(self.importWords, (word,))

    def replace_newline2br(self, string):
        return str.replace(string, '\n', '<br/>')

    def do_search_for_selection(self, wv):
        word = wv.page().selectedText()

        self.doWithinThread(word)

    @staticmethod
    def element_is_visible(xpath):
        try:
            ele = mydriver.find_element_by_xpath(xpath)
            return ele.is_displayed()
        except:
            return False

ImportWords()

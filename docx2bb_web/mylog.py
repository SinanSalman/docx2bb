import time


class LOG:
    def __init__(self):
        self.logtext = {'info':'', 'debug':''}
        self.TextWidth = 120

    def info(self,msg):
        self.debug(msg)
        if len(msg) > self.TextWidth:
            msg = msg[:self.TextWidth-3] + '...'
        if len(self.logtext['info']) > 0:
            self.logtext['info'] += '\n'
        self.logtext['info'] += msg

    def debug(self,msg):
        tm = time.strftime("%m-%d-%Y %H:%M:%S") + '|'
        if len(msg) > self.TextWidth-len(tm):
            msg = msg[:self.TextWidth-len(tm)-3] + '...'
        if len(self.logtext['debug']) > 0:
            self.logtext['debug'] += '\n'
        self.logtext['debug'] += tm + msg

    def clear(self,type):
        if type in ['info','debug']:
            self.logtext[type] = ''
        if type == 'all':
            for k,v in self.logtext.items():
                self.logtext[k] = ''

    def save(self,type,filename=''):
        if filename == '':
            return
        if type in ['info','debug']:
            with open(filename,'a') as logfile:
                logfile.write(self.logtext[type])

from functools import wraps
import win32com.client
import time

def dispatched_functions(function_decorator):
    def decorator(cls):
        for name, obj in vars(cls).items():
            if callable(obj):
                try:
                    obj = obj.__func__
                except AttributeError:
                    pass
                setattr(cls, name, function_decorator(obj))
        return cls
    return decorator

def on_call(func):
    @wraps(func)
    def wrapper(*args, **kw):
        #print('{} called'.format(func.__name__))
        try:
            res = func(*args, **{'funcName':func.__name__})
        finally:
            pass
            #print('{} finished'.format(func.__name__))
        return res
    return wrapper

@dispatched_functions(on_call)
class attachMate(object):
    def __init__(self,sessionPath='',timeOutValue=100000,*args, **kwargs):
        self.sessionPath=sessionPath
        self.system = win32com.client.Dispatch("EXTRA.System")
        self.system.TimeoutValue=timeOutValue
        self.ActiveSession=None
        self.Screen=None

    def getActiveSession(self,**kwargs):
        self.ActiveSession=self.system.ActiveSession
        return self.ActiveSession
        def Screen(self,**kwargs):
            return self.ActiveSession.Screen

    def Open(self,sessionPath,**kwargs):
        getattr(self.system.Sessions,kwargs['funcName'])(sessionPath)

    def OpenExtraSession(self,sessionPath,timeout=10,**kwargs):
        from os import system
        if '/' in sessionPath: sessionSplitedPath=sessionPath.split('/')
        elif '\\' in sessionPath: sessionSplitedPath=sessionSplitedPath.split('\\')
        sessionPathURL = sessionPath.replace(sessionSplitedPath[-1],'')
        from subprocess import Popen
        import time
        p = Popen(sessionSplitedPath[-1], shell=True,cwd=sessionPathURL)
        time.sleep(timeout)

    def Quit(self,**kwargs):
        '''
        Description:

            Closes all sessions and EXTRA! programs.
        '''
        getattr(self.system,kwargs['funcName'])()

    def Time(self,**kwargs):
        '''
        Description:

            Returns the current system time.
        '''
        return getattr(self.system ,kwargs['funcName'])()

    def ViewStatus(self,**kwargs):
        '''
        Description

            Starts the Status program.
        '''
        return getattr(self.system ,kwargs['funcName'])()


        

@dispatched_functions(on_call)
class extradriver(object):
    def __init__(self,attachMateSession, *args, **kwargs):
        self.sessionDriver=attachMateSession
        self.screenDriver=self.sessionDriver.Screen

    def Activate(self,**kwargs):
        '''
        Description:
            Makes the specified session the active window.
        
        Ex:
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.Activate()
        '''
        getattr(self.sessionDriver,kwargs['funcName'])

    ##def add

    def Area(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:
            Returns an Area object with the defined coordinates
        Ex:
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            area=xdriver.Area(2,47,3,60)
        '''
        return getattr(self.screenDriver,kwargs['funcName'])(StartRow,StartCol,EndRow,EndCol)

    def CaptureSetup(self,FileName='captured.log',Append=True,EnableCapture=True,StripEscSeq=True,**kwargs):
        '''
        Description:

            Determines how captured data will be written to a file.
        Ex:

            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.CaptureSetup()
            
        Optional params:

            Element         Description
            
            FileName        Sets the name of the file to which captured data will be written.
            Append          If the Boolean value appendBool is TRUE, appends data to the end of the file specified by FileName. If FALSE, overwrites any data previously stored in the file.
            EnableCapture   If the Boolean value enableBool is TRUE, allows data to be captured to file. If FALSE, disables data capture.
            StripEscSeq     If the Boolean value stripBool is TRUE, strips escape sequences from captured data before writing it to file.
        '''
        return getattr(self.sessionDriver,kwargs['funcName'])(FileName,Append,EnableCapture,StripEscSeq)

    ##Aply to waits
    ##def Clear(self,**kwarg):
    ##    pass

    def ClearComm(self,**kwargs):
        '''
        Description:

        Empties the communications buffers, and clears any outstanding conditions that could prevent transmitting or receiving data. It also cancels any escape sequence processing.

        Ex:

            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.ClearComm()
        '''
        getattr(self.sessionDriver,kwargs['funcName'])()

    def ClearHistory(self,**kwargs):
        '''
        Description:

        Removes all information from the History buffer.

        Ex:

            xdriver=extradriver(session)
            xdriver.ClearHistory()
        '''

        getattr(self.sessionDriver,kwargs['funcName'])()
    
    def ClearScreen(self,**kwargs):
        '''
        Description

            Clears display memory and sets all line attributes to normal.

        Ex:

            xdriver=extradriver(session)
            xdriver.ClearScreen()

        '''
        getattr(self.screenDriver,kwargs['funcName'])()

    def Close(self, **kwargs):
        '''
        Description:

            Closes the session.

        Ex:
        
            xdriver=extradriver(session)
            xdriver.Close()

        '''
        getattr(self.sessionDriver,kwargs['funcName'])()

    ## se aplica a sessions
    #def CloseAll(self, **kwargs):

    def CloseEx(self,options,**kwargs):
        '''
        Description:

            Closes the session.

        Example:

            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.CloseEx(1)

        Params:

            Use this value      To do this
            
            1                   Disconnect session without displaying a prompt
            4                   Save session while exiting
            8                   Exit session without saving
            16                  Prompt user to save session

        '''
        getattr(self.sessionDriver,kwargs['funcName'])(options)

    #Se aplica a sessions
    #def CloseAll(self,**kwargs):

    def ClosePrintJob(self,**kwargs):
        '''
        Description:

        Indicates the end of a print job, forces immediate printing of accumulated printer data from the print buffers, and sends a form feed to the print job, ejecting the current page.

        '''

        getattr(self.sessionDriver,kwargs['funcName'])()

    def Copy(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:

            Copies the select text to the Clipboard but leaves the selected text unchanged in the display.

        Ex:

            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.Copy(3,47,4,60)

        '''
        self.Area(StartRow,StartCol,EndRow,EndCol).Select()
        getattr(self.screenDriver,kwargs['funcName'])()

    def CopyAppend(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:

            Copies the selected text from a session to the existing contents of the Clipboard but leaves the selected text unchanged in the display.

        Ex:
        
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.CopyAppend(3,47,4,60)

        '''
        self.Area(StartRow,StartCol,EndRow,EndCol).Select()
        getattr(self.screenDriver,kwargs['funcName'])()

    def Cut(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:

            Removes selected text from the session and stores it in the Clipboard.


        Ex:
        
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.Cut(6,53,6,57)

        '''
        self.Area(StartRow,StartCol,EndRow,EndCol).Select()
        getattr(self.screenDriver,kwargs['funcName'])()

    def CutAppend(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:

            Removes selected text from a session and appends it to the existing contents of the Clipboard.

        Ex:
        
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.CutAppend(6,53,6,57)

        '''
        self.Area(StartRow,StartCol,EndRow,EndCol).Select()
        getattr(self.screenDriver,kwargs['funcName'])()


    def Delete(self,StartRow,StartCol,EndRow,EndCol,**kwargs):
        '''
        Description:

            Deletes the current selection.

        Ex:
        
            attachMate().OpenExtraSession('YourSessionPath/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.Delete(6,53,6,57)

        '''
        self.Area(StartRow,StartCol,EndRow,EndCol).Select()
        getattr(self.screenDriver,kwargs['funcName'])()

    def EnlargeFont(self,**kwargs):
        '''
        Description:

            Increases the session font size by one increment.

        '''
        getattr(self.sessionDriver,kwargs['funcName'])()

    def FieldAttribute(self,row,column,**kwargs):
        '''
        Description:

            Returns the field attribute value for a given row/column position on the current screen. Returns zero if an invalid row or column is provided, or if the current screen does not contain field formatting. Valid only for 3270 or 5250 emulation types.

        '''
        #TODO: Preguntar que significan estos atributos
        return getattr(self.screenDriver,kwargs['funcName'])(row,column)

    ## applies to systemsession object
    # def GetDirectory ()

    def GetString(self,row,column,length,**kwargs):
        '''
        Description:

            Returns the text from the specified screen location.

        Ex:

            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.GetString(6,53,25)
        
        '''
        initialRow=row
        initialColumn=column
        creatingCords=True
        column=column+length
        while creatingCords:
            if column>self.screenDriver.Cols:
                column=column-self.screenDriver.Cols
                row=row+1
            else: creatingCords=False
                        
        return self.Area(initialRow,initialColumn,row,column)

    #TODO: Se aplica sobre quickpads y toolbars
    # def HideAll(self,**kwargs):
    #    '''
    #     Description:

    #         Hides all visible QuickPad or Toolbar objects.




    #     '''

    def HostOptions(self,**kwargs):
        '''
        Description:
            Returns the HostOptions object associated with the session. Read-only.

        '''
        return getattr(self.screenDriver,kwargs['funcName'])

    def InvokeSettingsDialog(self,SettingsPage,SettingsTab,**kwargs):
        '''
        Description:

            Displays an EXTRA! Settings Dialog.

        Cons information: 
        
            http://docs.attachmate.com/extra/x-treme/apis/com/invokesettingsdialogmethod_con.htm

        
        '''
        return getattr(self.sessionDriver ,kwargs['funcName'])(SettingsPage,SettingsTab)

    ##Applye to multisessions and wwaits
    ##Def item method

    ##Applies to sessions
    ##def JumpNext()

    def MoveRelative(self,NumOfRows,NumOfcols,**kwargs):
        '''
        Description:
            Moves the cursor a specified number of rows and columns (or pages) from its current position.

        Ex:

            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.MoveRelative(0,35)

        '''
        getattr(self.screenDriver ,kwargs['funcName'])(NumOfRows,NumOfcols)

    def MoveTo(self,Row, Col,**kwargs):
        '''
        Description:

            Moves the cursor to the specified location.
            Certain VT hosts do not allow arbitrary cursor positioning, in which case this method will have no effect.

        Ex:
            
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.MoveTo(10,53)

        '''
        
        getattr(self.screenDriver ,kwargs['funcName'])(Row,Col)

    ##TODO: Preguntar por el screenName
    def NavigateTo(self,screenName,**kwargs):
        '''
        Description:

            Navigates to a specified host screen, recorded from a session window.

        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()

    def Paste(self,**kwargs):
        '''
        Description:

            For the Screen object, this method pastes Clipboard text at the current position or over the current selection.
        '''
        getattr(self.screenDriver ,kwargs['funcName'])()

    def PasteOn(self,Row,Col,**kwargs):
        '''
        Description:

            For the Screen object, this method pastes Clipboard text at the position.
        
        Ex:
    
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.PasteOn(10,53)

        '''
        self.MoveTo(Row,Col)
        self.Paste()

    def PasteContinue(self,**kwargs):
        '''
        Description

            Continues to insert more Clipboard data from the previous Paste operation.

        Ex:

            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.PasteContinue()

        '''
        getattr(self.screenDriver ,kwargs['funcName'])()

    def PrintDisplay(self,**kwargs):
        '''
        Description:

            Prints or saves on PDF format the current display screen.

        Ex:

            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.PrintDisplay()

        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()

    def PutString(self,String,Row,Col,**kwargs):
        '''
        Description:

            Puts text in the specified location on the screen.

        Ex:
            
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.PutString('Juan',6,53)
        '''
        getattr(self.screenDriver ,kwargs['funcName'])(String,Row,Col)

    ##TODO: pendiente de utilidad
    ##def ReceiveFile

    def ReduceFont(self,**kwargs):
        '''
        Description:
            Reduces the session font size by one increment.

        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()

    ##Applies to object Waits
    ##def Remove

    ##RemoveAll

    def Reset(self,**kwargs):
        '''
        Description:

            Used with the Session object, Reset returns the display to its power-up operating state. Session.Reset is only valid for VT terminal sessions.
        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()

    ##def ResetAttributes
    ##def ResetColors
    ##def ResetTabs

    def Save(self,**kwargs):
        '''
        Description:

            Saves the current settings of the session.

        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()
    
    def SaveAs(self,FileName,**kwargs):
        '''
        Description:

            Saves the current settings of the session.
            
        '''
        getattr(self.sessionDriver ,kwargs['funcName'])(FileName)

    def Search(self,Text,**kwargs):
        '''
        Description:

            Returns an Python dict object with the cords of specified object in the search.

        Ex:
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            print(xdriver.Search('User  . . . . . . . . . . . . . .'))

        '''
        ExtraObject=getattr(self.screenDriver ,kwargs['funcName'])(Text)
        ExtraObjectCords = {
            'Bottom': ExtraObject.Bottom,
            'Left': ExtraObject.Left,
            'Right': ExtraObject.Right,
            'Top': ExtraObject.Top
        }

        return ExtraObjectCords


    def Select(self,StartRowBottom,StartColLeft,EndRowTop,EndColRight,**kwargs):
        '''
        Description:

            Selects data in an Area based on cords.

        Ex:
            
            attachMate().OpenExtraSession('C:/Users/juan.restrepo/extraDriverProject/as400-demostracion.edp')
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.Select(6,17,6,49)

        '''
        self.Area(StartRowBottom,StartColLeft,EndRowTop,EndColRight).Select()

    def SelectExtraObject(self,ExtraObject,**kwargs):
        '''
        Description:

            Selects the ExtraObject(Dict)

        Ex:
            
            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            user=xdriver.Search('User')
            xdriver.SelectExtraObject(user)

        '''
        Left=ExtraObject['Left']
        Bottom=ExtraObject['Bottom']
        Right=ExtraObject['Right']
        Top=ExtraObject['Top']
        self.Select(Bottom,Left,Top,Right)

    def SelectAll(self,**kwargs):
        '''
        Description:

            Selects the entire screen and returns an Area object.
        '''
        getattr(self.screenDriver ,kwargs['funcName'])()


    ##TODO: hacer clase ExtraObject

    #TODO: preguntar para que sirve esto
    #def SendFile

    def SendInput(self,Text,**kwargs):
        '''
        Description

            Sends the specified text to the Screen object, simulating incoming data from the host.
        '''
        getattr(self.screenDriver ,kwargs['funcName'])(Text)
    
    def SendKeys(self,Text,**kwargs):
        '''
        Description

            Sends the specified text to the Screen object, simulating incoming data from the host.

        Ex:

            session=attachMate().getActiveSession()
            xdriver=extradriver(session)
            xdriver.SendKeys('qweqwe')

        '''
        getattr(self.screenDriver ,kwargs['funcName'])(Text)

    def TransferFile(self,**kwargs):
        '''
        Description:

            Displays the File Transfer dialog box.
        '''
        getattr(self.sessionDriver ,kwargs['funcName'])()

    #TODO: se aplica al host options
    # def UDKClear(self,**kwargs):
    #     '''
    #     Description:

    #         Clear all the user-defined keys value.
        
    #     '''

    def UpdateStatusBar(self,String,**kwargs):
        '''
        Description

            Displays the specified string in the session's status bar.
        '''
        getattr(self.sessionDriver ,kwargs['funcName'])(String)

    def writeOn(self,ExtraObject,text,**kwargs):
        objectCords=self.Search(ExtraObject)
        self.SelectExtraObject(objectCords)
        col=objectCords['Right']
        row=objectCords['Top']
        editablerow,editablecolumn=self.searchNextEditableFields(row,col)
        self.PutString(text,editablerow,editablecolumn)

    def searchNextEditableFields(self,startrow,startcol,**kwargs):
        for i in range(startcol,80):
            if self.FieldAttribute(startrow,i)==192:
                finalrow=startrow
                finalcolumn=i+1
                return finalrow,finalcolumn
        return 1,1

    def GoToNextScreen(self,**kwargs):
        self.SendKeys('<Enter>')
        while(self.screenDriver.Updated): time.sleep(0.2)

if __name__=='__main__':

    screenLoggin = {
        'User':' User  . . . . . . ',
        'Password':'Password  . . . . . . . .',
        'Program':'Program/procedure . . . . ',
        'Menu':'Menu',
        'Current':'Current library'
    } 

    As400MainMenu = {
        'Command':'==>',
        }

    Emulator=attachMate()
    Emulator.OpenExtraSession('../as400-demostracion.edp')
    session=attachMate().getActiveSession()
    xdriver=extradriver(session)
    xdriver.writeOn(screenLoggin['User'],'Juan')
    xdriver.writeOn(screenLoggin['Password'],'DEMO5250')
    xdriver.writeOn(screenLoggin['Program'],'no idea') 
    xdriver.writeOn(screenLoggin['Menu'],'inventado')
    xdriver.writeOn(screenLoggin['Current'],'actual')
    xdriver.GoToNextScreen()
    xdriver.writeOn(As400MainMenu['Command'],'comando')   
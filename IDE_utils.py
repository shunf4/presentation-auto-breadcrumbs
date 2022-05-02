#!/usr/bin/python
# -*- coding: utf-8 -*-

""" Run/Debug LibreOffice Python macros in IDE [#IDE]_

    This module provides:
    i.   a substitute to `XSCRIPTCONTEXT` (Libre|Open)Office built-in,
         to be used within `Python IDE`_ such as Anaconda, Geany,
         KDevelop, PyCharm, etc..
    ii.  a (Libre|Open)`Runner` context manager with
       - `start`, `stop` paradigms to launch *Office instances
         and to facilitate `setup`, `tearDown` unit testing steps 

    Instructions:

    1.   Copy this module into your <OFFICE>/program/ directory
      OR Include it into your IDE project directory
    2.   Include one of the below examples into your Python macro
    3.   Run your (Libre|Open) macro from your preferred IDE

    Examples:

    import uno
    def my_1st_macro(): pass  # Your code goes here
    def my_2nd_macro(): pass  # Your code goes here
    def my_own_macro(): pass  # Your code goes here

    g_exportedScripts = my_1st_macro, my_2nd_macro, my_own_macro

    # i. Runners.json argument file /OR/ (Libre|Open)Office pipe
    if __name__ == "__main__":
        import IDE_utils as ide
        with ide.Runner() as jesse_owens:  # Start, Stop
            XSCRIPTCONTEXT = ide.XSCRIPTCONTEXT  # Connect, Adapt
            my_1st_macro()  # Run

    # ii. {pgm: [accept, *options]} service-options pair(s)
    if __name__ == "__main__":
        import IDE_utils as geany
        pgm = {'/Applications/LibreOffice.app/Contents/MacOS/soffice':
               ['--accept=pipe,name=LinusTorvalds;urp;',
                '--headless', '--nodefault', '--nologo']}
        with geany.Runner() as carl_lewis:  # Start, Stop
            XSCRIPTCONTEXT = geany.XSCRIPTCONTEXT  # Connect, Adapt
            my_2nd_macro()  # Run

    # iii. Named pipe bridge
    if __name__ == "__main__":
        from IDE_utils import connect, ScriptContext
        ctx = connect(pipe='LinusTorvalds')  # Connect
        XSCRIPTCONTEXT = ScriptContext(ctx)  # Adapt
        my_own_macro()  # Run

    Imports:
        itertools - retry decorator
        json - services' running conditions
        logging
        officehelper - bootstrap *Office
        os - Check file
        re - Parse UNO-URL's
        subprocess - Control *Office services
        sys - Identify platform
        time - sleep
        traceback
        uno

    Interfaces:
        com.sun.star.script.provider.XSCRIPTCONTEXT

    Exceptions:
        BootstrapException - from `officehelper`
        NoConnectException - in `ScriptContext`
        NotImplementedError - in `ScriptContext`
        OSError, RuntimeError - in `Runner`
        RuntimeException - in stop

    Classes:
        `Runner(soffice=None)` - Start, stop *Office services
        `ScriptContext(ctx)` - Implement XSCRIPTCONTEXT

    Functions:
        `connect(host="localhost", port=2002, pipe=None)`
        `start(soffice=None)` - Start *Office services
        `stop` - Stop *Office services
        `killall_soffice` - Interrupt `soffice` running tasks

    see also::
        `help(officehelper)`

    warning:: Only `soffice` binaries get processed.
        Tested platforms are Linux, MacOS X & Windows

    Created on: Dec-2017
    Version: 0.8
    Author: LibreOfficiant
    Acknowledgements:
      - Kim Kulak, for his Python bootstrap.
      - Christopher Bourez, for his tutorial.
      - Tsutomu Uchino (Hanya), for Interface Injection first implementation.
      - Mitch Frazier, for inspiring start/stop, on-demand options and pooling.
      - ActiveState, for retry python decorator
      - Joerg Budichewski, for PyUNO.

    :References:
    .. _Python IDE: https://wiki.documentfoundation.org/Macros/Python_Basics
    .. [#IDE] "Integrated Development Environment"
"""
from __future__ import print_function

import atexit, \
    itertools, \
    json, \
    logging, \
    officehelper, \
    os, \
    re, \
    subprocess, \
    sys, \
    time, \
    traceback, \
    uno

_INFO, _DEBUG, _EMULATE_OFFICEHELPER = False, False, True
logging.basicConfig(format='%(asctime)s %(levelname)8s %(message)s')
if _INFO: logging.getLogger().setLevel(logging.INFO)
if _DEBUG: logging.getLogger().setLevel(logging.DEBUG)

#from uno import RuntimeException
from com.sun.star.lang import DisposedException
from com.sun.star.script.provider import XScriptContext
from com.sun.star.connection import NoConnectException
from officehelper import BootstrapException


RUNNERS = 'Runners.json'
_SECONDS = 3

if "UNO_PATH" in os.environ: logging.debug(os.environ["UNO_PATH"])

def create_service():
    ''' Build a JSON config.

        {pgm, [accept, *options]} dict pairs to be started/stopped as
        instances of (Libre|Open)Office. The created JSON file can be
        PRETTYfied using " python -m json.tool "

    '''
    _MY_PIPE = 'LibreOffice'
    _MY_PORT = 8100  # 2002

    _MAC_libO = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    _LNX_ANY = 'soffice'
    _LNX_OO4 = '/opt/openoffice4/program/soffice'
    _WIN_libO = 'C:\Program Files\LibreOffice 5\program\soffice.exe'
    _WIN_libO_X86 = _WIN_libO.replace('Program Files','Program Files (x86)')
    _WIN_OOo_USB = 'USB:\PortableApps\aOO-4.1.5\App\openoffice\program\soffice.exe'

    _libO_FOREGROUND = ['--accept="pipe,name='+_MY_PIPE+';urp;"'
                       ]  # LibreOffice - Foreground - visible instance
    
    _libO_BACKGROUND = ['--accept="socket,host=localhost,port='+str(_MY_PORT)+';urp;"',
                       '--headless',
                       #'--invisible',
                       '--minimized',
                       '--nodefault',
                       '--nologo',
                       '--norestore',
                       #'--safe-mode',
                       #'--unaccept={UNO-URL}',
                       #'--unaccept=all',
                       '--language=fr'
                       ]  # LibreOffice - Background - localhost,port#
    _aOO_BACKGROUND = ['-accept="socket,host=localhost,port=2002;urp;"',
                      '-headless',
                       #'-invisible',
                       '-maximized',
                       '-minimized',
                      '-nodefault',
                      '-nologo',
                      '-norestore'
                      ]  # OpenOffice - Background - localhost,port#

    services = {_MAC_libO: _libO_BACKGROUND,
                _LNX_OO4: _aOO_BACKGROUND,
                _WIN_libO: _libO_BACKGROUND,
                _WIN_libO+' ': _libO_FOREGROUND,
                _WIN_OOo_USB: _aOO_BACKGROUND
                }  # concurrent (Libre|Open)Office instances

    with open(RUNNERS, 'w') as f:
        json.dump(services, f)


class Runner():  # (Libre|Open)Office Runner
    """ (Libre|Open)`Runner` context manager

    o  It holds `start`, `stop` paradigms to launch `sOffice` instances &
    o  It facilitates `setup`, `tearDown` unit testing steps

    Description:
        Starts, stops zero-to-many (Libre|Open)Office processes
        according to an optional JSON file or argument containing
        {pgm: [accept, *options]} key-values service pairs.

    Recommendation:
    o  Concurrent instances/services require that --accept UNO-Urls are
       unique. In others words (host, port#) sockets and named (pipe)
       must not be redefined.

    Examples:

    import IDE_utils as ide
    with ide.Runner() as usain_bolt:  # Starts/stops 'soffice' instances
        XSCRIPTCONTEXT = ide.XSCRIPTCONTEXT
        # Your code goes here

    import IDE_utils as ide
    task = {'D:\Portable\App\openoffice\program\soffice.exe':
            ['-accept="pipe,name=OpenOffice;urp;"'
             ]  # OpenOffice - Foreground - visible instance
            }
    with ide.Runner(soffice=task) as carl_lewis:  # Portable OpenOffice
        XSCRIPTCONTEXT = ide.XSCRIPTCONTEXT
        # Your code goes here

    from IDE_utils import start, stop, XSCRIPTCONTEXT
    try:
        start()  # starts ALL 'soffice' JSON filed pgms
        # Your code goes here
    finally:
        stop()  # interrupts ALL 'soffice' instances


    see also: `XSCRIPTCONTEXT` built-in & `ScriptContext`
    """

    def __init__(self, soffice=None):
        self.services = {}  # (pgm, serv_descr) pairs
        self.processes = {}  # (uno-url, process) key-value pairs
        self.pool = {}  # Context pool, cf. ScriptContext.pool
        if soffice is None or type(soffice) != dict:
            logging.debug("READing.. default JSON file services' list")
            self.services = Runner._read_service()
        else:
            logging.debug("READing.. JSON argument services' list")
            self.services = soffice
    def __enter__(self):
        logging.debug("ENTERing.. Runner context manager")
        return self._start()
    def __exit__(self, exctype, exc, tb):
        logging.debug("EXITing.. Runner context manager")
        if tb is None:  # no traceback means no error/exception
            self._stop()
        else:
            pass  # Need to review this !!
        return

    @staticmethod
    def _accept2Uno(accept_url):
        """ Convert the --accept connection string into an UNO-URL

        :param accept_url: accept connection string '--accept..'
        :return UNO-URL:
        :rtype: str

        >>> Runner._accept2Uno('--accept="socket,host=localhost,port=2002;urp;"')
        'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext'

        >>> Runner._accept2Uno('-accept="socket,host=127.0.0.1,port=8001;urp;"')
        'uno:socket,host=127.0.0.1,port=8001;urp;StarOffice.ComponentContext'

        >>> Runner._accept2Uno('--accept="pipe,name=LinusTorvalds;urp;"')
        'uno:pipe,name=LinusTorvalds;urp;StarOffice.ComponentContext'

        >>> Runner._accept2Uno('--accept=LibreOffice UNO-URL;')
        'uno:LibreOffice UNO-URL;StarOffice.ComponentContext'

        >>> Runner._accept2Uno("-accept=OpenOffice -accept string;")
        'uno:OpenOffice -accept string;StarOffice.ComponentContext'

        """
        if type(accept_url) != str: return None
        pattern = '-{1,2}accept='
        if re.search(pattern, accept_url):
            uno_url = re.sub(pattern,'uno:',accept_url)  # remove prefix
            uno_url = ''.join([uno_url.replace('"', ''),  # remove quotes
                'StarOffice.ComponentContext'])  # append Obj-name
            return uno_url
    @staticmethod
    def _isOfficeBinary(bin_path):
        """ Check input for expected (Libre|Open)Office binary

        :param bin_path: program location e.g. '<OFFICE_PATH>soffice'
        :type bin_path: str
        :rtype: boolean

        >>> Runner._isOfficeBinary(3644)
        False

        >>> Runner._isOfficeBinary("xyzsoffice1234509876")
        True

        >>> Runner._isOfficeBinary('soffice')
        True
        """
        if type(bin_path) != str: return False
        return True if re.search('soffice', bin_path) else False
    @staticmethod
    def _read_service():
        services = {}  # (pgm, options) key-value pairs cf. _create_config()
        if os.path.isfile(RUNNERS):
            with open(RUNNERS, 'r') as f:
                services = json.load(f)  # (pgm, [ accept, *options]) pairs
        return services
    def connect(self, host=None, port=None, pipe=None, uno_url=None):
        ctx = ScriptContext._connect(self.pool, host=None,
            port=None, pipe=None, uno_url=uno_url, flush=True)
        return ctx
    def _start(self):
        logging.info('STARTing (Libre|Open)Office instances..')
        for pgm, options in self.services.items():
            if not Runner._isOfficeBinary(pgm): continue
            try:
                cmd = pgm+' '.join(self.services[pgm])
                key = self._accept2Uno(options[0])
                options.insert(0, pgm)
                self.processes[key] = subprocess.Popen(options).pid
                logging.debug('STARTed.. %s ' % cmd)
                officehelper.sleep(_SECONDS)  # No longer required ?
                self.pool[key] = self.connect(uno_url=key)
                XSCRIPTCONTEXT = self.connect(uno_url=key)
            except OSError as e:  # WindowsError super class
                ''' OSError is OS-agnostic '''
                logging.error(pgm + " not found.")
        logging.debug(self.processes)
        logging.debug('STARTed (Libre|Open)Office instances..')
    def _stop(self):
        logging.info('STOPping (Libre|Open)Office instances..')
        if len(self.pool) != 0:
            _terminate_desktops(self.pool)


def start(soffice=None):
    """ START (Libre|Open)sOffice instances """
    Runner(soffice=soffice)._start()

_DELAYS = [(0, 1, 5, 30, 180, 600, 3600),  # try 0, 1, 5, 10, .. sec.
           [0] + [0.5] * 19,  # try 20 times each half-second.
           (0, 1, 1, 1, 1, 1)]  # try 6 times each second.
CONNECT_DELAYS = _DELAYS[0]
CONNECT_EXCEPTIONS = NoConnectException
#CONNECT_REPORT = print  # How about logging.info instead ?
CONNECT_REPORT = lambda *args: None  # Silent connections

# Credit:
# http://code.activestate.com/recipes/580745-retry-decorator-in-python/
def retry(delays=(0, 1, 5, 30, 180, 600, 3600),
          exception=Exception,
          report=lambda *args: None):
    ''' Decorator: Retry certain steps which may fail sometimes '''
    def wrapper(function):
        def wrapped(*args, **kwargs):
            problems = []
            for delay in itertools.chain(delays, [ None ]):
                try:
                    return function(*args, **kwargs)
                except exception as problem:
                    problems.append(problem)
                    if delay is None:
                        report("\n retryable failed definitely:", problems)
                        raise
                    else:
                        report("retryable failed:", problem,
                            "-- delaying for %ds" % delay)
                        time.sleep(delay)
        return wrapped
    return wrapper


class ScriptContext(XScriptContext):
    """ Substitute (Libre|Open)Office XSCRIPTCONTEXT built-in

    Can be used in IDEs such as Anaconda, Geany, KDevelop, PyCharm..
    in order to run/debug Python macros.

    Implements: com.sun.star.script.provider.XScriptContext

    Usage:

    ctx = connect(pipe='RichardMStalman')
    XSCRIPTCONTEXT = ScriptContext(ctx)

    ctx = connect(host='localhost',port=1515)
    XSCRIPTCONTEXT = ScriptContext(ctx)

    see also: `Runner`
    """
    '''
    cf. <OFFICEPATH>/program/pythonscript.py
    cf. https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=53748
    https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.xhtml?search=8100
    '''
    def __init__(self, ctx):
        self.ctx = ctx
    def getComponentContext(self):
        return self.ctx
    def getDesktop(self):
        return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
    def getDocument(self):
        return self.getDesktop().getCurrentComponent()
    def getInvocationContext(self):
        raise os.NotImplementedError
    @staticmethod
    def _connect(ctx_pool, host='localhost', port=2002, pipe=None,
                 uno_url=None, flush=True):
        ''' (re)Connect to socket/pipe *Office instances or Fail

        arguments:
        ctx_pool: {key: ctx} pool of ComponentContext to explore/feed
        
        '''
        if uno_url:
            if str(uno_url) in ctx_pool:
                return ctx_pool[uno_url]
            # ~ else:
                # ~ uno_url = uno_url
        elif pipe:
            if str(pipe) in ctx_pool:
                return ctx_pool[pipe]
            else:
                uno_url = ''.join(['uno:pipe,name=',
                    str(pipe),
                    ';urp;StarOffice.ComponentContext'])
        else:
            if int(port) in ctx_pool:
                return ctx_pool[port]
            else:
                uno_url = ''.join(['uno:socket,host=',host,',port=',str(port)])
                uno_url = ''.join([uno_url,';urp;StarOffice.ComponentContext'])

        localContext = uno.getComponentContext()
        resolver = localContext.getServiceManager().createInstanceWithContext(
                        "com.sun.star.bridge.UnoUrlResolver", localContext )

        logging.info('CONNECTing to ' + uno_url)
        @retry(delays=CONNECT_DELAYS,
               exception=CONNECT_EXCEPTIONS,
               report=CONNECT_REPORT)
        def resolve():
            return resolver.resolve(uno_url)
        ctx = resolve()  # Raises NoConnectException

        if not flush: return ctx  # otherwise pool contexts to kill
        if uno_url:
            ctx_pool[uno_url] = ctx
        elif pipe:
            ctx_pool[pipe] = ctx
        else:
            ctx_pool[(host, port)] = ctx
        return ctx


ScriptContext.pool = {}
""" keys categorize in 3 types:
    - uno-urls connection strings
    - socket (host,port) tuples
    - pipe names
"""
'''NOTE: Runner() objects also hold a component context pool '''

def connect(host='localhost', port=2002, pipe=None, flush=False):
    ''' Connect to socket/pipe *Office instances or Fail

    Keyword arguments:
    host: 'localhost' or IP address
    port: socket #
    flush: Whether to force service termination ( default is False )

    return: uno.getComponentContext() service equivalent

    raises:
    NoConnectException - Unreachable service
    ConnectionSetupException - malformed uno_url
    '''
    return ScriptContext._connect(ScriptContext.pool, host=host,
        port=port, pipe=pipe, uno_url=None, flush=flush)


# ============
#  INITIALIZE
# ============

def _bootstrap():
    ''' Initialize a default piped service '''
    # soffice script used on *ix, Mac; soffice.exe used on Win
    # --startup LibreOffice options; -startup OpenOffice options  
    if "UNO_PATH" in os.environ:
        sOffice = os.environ["UNO_PATH"]
    else:
        sOffice = "" # hope for the best

    sOffice = os.path.join(sOffice, "soffice")
    if sys.platform.startswith("win"):
        sOffice += ".exe"

    options = ['-accept=pipe,name=OfficeHelper;urp;', 
        '-nodefault', '-nologo']
    if sys.version_info.major == 3:  # OpenOffice uses 2.x
        options = ['--accept=pipe,name=OfficeHelper;urp;',
            '--nodefault', '--nologo']

    oh = {sOffice: options}
    start(soffice=oh)
    ctx = connect(pipe='OfficeHelper')

    return ctx

logging.info('BOOTSTRAPping (Libre|Open)Office instance..')
if _EMULATE_OFFICEHELPER:
    _ctx = _bootstrap()
else:
    _ctx = officehelper.bootstrap()

ScriptContext.pool['officehelper'] = _ctx # Force service termination
XSCRIPTCONTEXT = ScriptContext(_ctx)
''' Substitute XSCRIPTCONTEXT built-in '''

# ===========
#  TERMINATE
# ===========

# atexit.register(stop)
''' Make sure services are released 

    Whenever ScriptContext() is used but Runner() isn't, all services
    that were connected to, including officehelper random pipe, have to
    be terminated using stop() routine.

'''  # STOP `officehelper` (Libre|Open)Office instances
# @atexit.register
def stop():
    """ STOP all (Libre|Open)sOffice instances """
    logging.info('EXITing '+__name__)
    try:
        if ScriptContext.pool:  # non-empty pool
            _terminate_desktops(ScriptContext.pool)
    except (DisposedException) as e:
        ''' URP bridge already released '''
        logging.error(e)

def _terminate_desktops(ctx_pool):
    ''' Stop (Libre|Open)Office active sessions

    Two different ComponentContext pools land here:
    - module level pool
    - Runner() class level pool
    '''  #  Process trees are different between LibreOffice/OpenOffice
    #  . LibreOffice: Individual process tree / service
    #  . AOpenOffice: Separate process trees /service
    #  . OpenOffice: Single process tree / services

    logging.debug(ctx_pool.keys())
    for ctx in ctx_pool.values():  # {key: ctx} pairs
        ScriptContext(ctx).getDesktop().terminate()
    if _DEBUG: killall_soffice()
    # A bit drastic, but helps detecting malfunctions

def killall_soffice():
    ''' kill all pending `soffice` instances '''
    platform = officehelper.platform
    if platform.startswith('linux') or platform == 'darwin':
        subprocess.Popen(['killall','soffice'])
        #~ subprocess.run(['killall','soffice'])
    elif platform.startswith('win'):
        subprocess.Popen(['taskkill','/f','/t', '/im','soffice.exe'])
        #~ subprocess.run(['taskkill','/f','/t', '/im','soffice.exe'])
    else:
        raise RuntimeError(' Unsupported {} platform'.format(platform))


# ======
#  TEST
# ======

if __name__ == "__main__":

    #~ import sys
    #~ print(sys.path)  # Detect portable aOO-4.1.5 <PATH>

    #~ create_service()

    import doctest, sys
    #~ logging.info('DOCTESTing.. %s \n' % sys.modules[__name__])
    __test__ = {
                #~ "Class:Runner:readConfig": "_read_service",
                #~ "Class:ScriptContext:connect": "_connect",
                "Function:terminateDesktops": "_terminate_desktops"
                }  # Additional DocTests for private attributes
    doctest.testmod()

    # import IDE_utils as me
    # help(me)
    # help(Runner)
    # help(_bootstrap)

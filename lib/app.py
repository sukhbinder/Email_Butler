import os
import sched
import shlex
import subprocess
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
from dateutil import parser

VBSCRIPT = """
Const olFolderCalendar = 9
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set f = fso.OpenTextFile("{fname}", 2)

Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar)
Set colItems = objFolder.Items
colItems.Sort("[Start]")
colItems.IncludeRecurrences = "True"
strFilter = "[Start] >= '{stdate}' AND [Start] <= '{enddate}'"
Set colFilteredItems = colItems.Restrict(strFilter)

For Each objItem In colFilteredItems
    If objItem.Start > Now Then
        f.WriteLine  objItem.Start   & "," & objItem.Subject  & ","  & objItem.Duration & ","  & objItem.Location
    End If
Next

f.Close
"""

APPCMD = "cscript //Nologo '{fname}'"

event_schedule = sched.scheduler(time.time, time.sleep)


def write_script(begin=datetime.today(), days=1):
    td = tempfile.gettempdir()
    file1 = Path(td) / "appointment.txt"
    file2 = Path(td) / "_app.vbs"

    file1.touch()
    endday = begin+timedelta(days=days)

    endday = endday.strftime("%m/%d/%Y")
    begin = begin.strftime("%m/%d/%Y")
    with open(file2, "w") as fout:
        fout.write(VBSCRIPT.format(fname=file1, stdate=begin, enddate=endday))
    return file2, file1


def run_script(scriptfile):
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    apptxt = APPCMD.format(fname=scriptfile)
    cmds = shlex.split(apptxt)
    iret = subprocess.call(cmds, startupinfo=startupinfo)
    return iret


def process_time(output):
    MIN5 = -5*60
    MIN15 = -15*60
    MINP5 = 5*60
    with open(output, "r") as f_in:
        lines = f_in.readlines()

    now = datetime.now()
    dates = np.array([(now - parser.parse(item.split(",")[0])
                       ).total_seconds() for item in lines])
    What = np.array([item.split(",")[1] for item in lines])
    duration = np.array([item.split(",")[2] for item in lines])
    fmintasks = dates > MIN5
    fifmintasks = dates > MIN15
    return (What[fmintasks], What[fifmintasks])


def get_outlook_schedule(begin=datetime.today(), days=1, show=False):
    if isinstance(begin, str):
        begin_d = parser.parse(begin)
    else:
        begin_d = begin

    script, output = write_script(begin=begin_d, days=days)
    iret = run_script(script)

    if show:
        os.startfile(output)
    return output


def scheduler_ontime(output):
    with open(output) as fin:
        lines = fin.readlines()

    for line in lines:
        t, story, *rest  = line.split(",")
        tt = parser.parse(t)
        eve = sched.Event(tt.timestamp(), 1, announce, (story, "zero"), {})
        if eve not in event_schedule.queue:
            event_schedule.enterabs(
                tt.timestamp(), 1, announce, (story, "zero"))


def subprocess_say(msg):
    startinfo = subprocess.STARTUPINFO()
    startinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    iret = subprocess.run(["say", msg],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                          startupinfo=startinfo)
    return iret


def announce(task, mins="couple of "):
    msg = "Meeting titled {0} in {1} minutes".format(task, mins)
    iret = subprocess_say(msg)


def mainrun():
    script, output = write_script()
    iret = run_script(script)
    # os.startfile(output)
    scheduler_ontime(output)
    five_min, fifteen_min = process_time(output)
    for task in fifteen_min:
        if task in five_min:
            announce(task)
        else:
            announce(task, mins="few")

    for task in five_min:
        announce(task)

    event_schedule.enter(360, 1, mainrun, ())


def main():
    event_schedule.enter(30, 1, mainrun, ())
    event_schedule.run()


if __name__ == "__main__":
    main()

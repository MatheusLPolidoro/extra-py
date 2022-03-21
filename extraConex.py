from win32com import client

settletime = 1000

system = client.GetObject('Reflection Workspace')
screen = system.GetObject("Frame").view(1).Control.Screen
screen.putText2('X', 5, 14)
screen.Wait(settletime)
screen.putText2('1', 5, 14)


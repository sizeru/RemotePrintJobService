﻿sc create RemotePrintJobService binPath= "C:\Program Files\RemotePrintJobService\RemotePrintJobService.exe" DisplayName= "Remote Print Job Service" start= auto
sc description RemotePrintJobService "Fetches And Prints Remote Print Jobs"
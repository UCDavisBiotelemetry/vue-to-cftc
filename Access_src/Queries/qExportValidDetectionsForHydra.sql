SELECT Format(DateAdd("h",8,DetectDate),"yyyy-mm-dd hh:nn:ss") AS [Date and Time (UTC)], SWITCH(VR2SN<300000,"VR2W-",VR2SN>540000,"VR2AR-",(VR2SN<450050) AND (VR2SN>=450000),"VR2C180-",VR2SN>=450000,"VR2C69-",VR2SN>=300000,"VR2W180-") & CStr(VR2SN) AS Receiver, Codespace & "-" & CStr(TagID) AS Transmitter, Data AS [Sensor Value], Units AS [Sensor Unit]
FROM Import_Detections;

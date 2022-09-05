using System;
using System.Collections.Generic;
using System.Text;

namespace RaceDisplay
{

    //purpose of this class is to allow the user to update the information on the text
	//boxes from the from the commserver object.  This class will take care of the destruction of
	//all objects created by the commserver by calling the closeall procedure in the commserver object

    class CommSvr
    {

        internal delegate void CommEventMessage(string strMessage);
        public event CommEventMessage CommEventMsg;
        public event CommEventMessage CommEventTimerMsg;

        internal delegate void CommEventObjects(short shrRaceNumber);
        public event CommEventObjects CommEventNewOdds;
        public event CommEventObjects CommEventRaceChange;
        public event CommEventObjects CommEventNewRO;

        public RaceFxCommSvr.clsCommSvr oCommServer;
        //public bool p_blnEventBusy;

        public CommSvr()    //Initialize
        {
            oCommServer = new RaceFxCommSvr.clsCommSvr();
            if (oCommServer != null)
            {
                oCommServer.Message += new RaceFxCommSvr.__clsCommSvr_MessageEventHandler(oCommServer_Message);
                oCommServer.TimerMsg += new RaceFxCommSvr.__clsCommSvr_TimerMsgEventHandler(oCommServer_TimerMsg);
                oCommServer.NewOdds += new RaceFxCommSvr.__clsCommSvr_NewOddsEventHandler(oCommServer_NewOdds);
                oCommServer.RaceChange += new RaceFxCommSvr.__clsCommSvr_RaceChangeEventHandler(oCommServer_RaceChange);
                oCommServer.NewRO += new RaceFxCommSvr.__clsCommSvr_NewROEventHandler(oCommServer_NewRO);
            }
            
        }

        ~CommSvr()
        {
            try
            {
                oCommServer.CloseAll();
                oCommServer = null;
            }
            catch(Exception ex)
            {
                ex.ToString();
            }
        }

        void oCommServer_Message(ref string strMessage)
        {
            //when the event for the tote comm object is raised then
            //we write to the text boxes
            CommEventMsg(strMessage);
        }

        void oCommServer_TimerMsg(ref string strTimer_Msg)
        {
            CommEventTimerMsg(strTimer_Msg);
        }

        void oCommServer_NewOdds(ref short intRace)
        {
            CommEventNewOdds(intRace);
        }

        void oCommServer_RaceChange(ref short intRace)
        {
            CommEventRaceChange(intRace);
        }

        void oCommServer_NewRO(ref short intRace)
        {
            CommEventNewRO(intRace);
        }

    }
}
	
//    Private Sub oCommServer_ResultsOfficial(ByRef intRace As Short) Handles oCommServer.ResultsOfficial
//        frmMain.lblInfoRaceResultsOff.Text = "Race " & intRace
//    End Sub
	
//    Private Sub oCommServer_StartTimerClock(ByRef blnStartClock As Boolean) Handles oCommServer.StartTimerClock
//        '  If frmTimer.Visible Then
//        '    frmTimer.StartClock
//        '  End If
//    End Sub
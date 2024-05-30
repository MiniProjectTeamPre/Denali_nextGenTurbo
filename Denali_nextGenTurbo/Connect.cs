using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//Adding visa liberary reference
using Ivi.Visa.Interop;

namespace Denali_nextGenTurbo {
    class Connect
    {
        Ivi.Visa.Interop.ResourceManager mRM = null;
        //Ivi.Visa.Interop.FormattedIO488Class mFIC = null;
        Ivi.Visa.Interop.FormattedIO488 mFIC = null;

        public Connect()
        {
            // Object of Interface
            mRM = new Ivi.Visa.Interop.ResourceManager();
            //mFIC = new Ivi.Visa.Interop.FormattedIO488Class();
            mFIC = new Ivi.Visa.Interop.FormattedIO488();
        }

        /* This Function Opens Connection to the Instrument using VISA Interface 
         * for the address entered by user at the time of Execution. */
        public string ConnectInstr(string instAddress)
        {
            //Check for Exception
            try
            {
                mFIC.IO = (IMessage)mRM.Open(instAddress, AccessMode.NO_LOCK,0,"");



                //mFIC.IO = (Ivi.Visa.Interop.IMessage)mRM.Open(instAddress, Ivi.Visa.Interop.AccessMode.NO_LOCK, 2000, "Timeout = 20000 ; TerminationCharacter = 10 ; TerminationCharacterEnabled=true");
                // Return string if connection to Instrument is successful
                return "Connected";
            }
            catch (Exception oExp)
            {
                // Return Exception message if not able to connect with Instrument
                return oExp.Message.ToString();
            }
        }

        /* This Function Closes the Connection to the Instrument that was opened by the user*/
        public bool DisConnectInstr()
        {
            try
            {
                // To close VISA Connection.
                (mFIC.IO).Close();
                return true;
            }
            catch (Exception oExp)
            {
                return false;
            }
        }


        /* ****************************************************************
         Send the SCPI command to the Instrument using VISA connection.
         And if SCPI is sent successful, this function returns "TRUE" 
         Else (if there is any exception) returns "FALSE"
         ******************************************************************/

        public bool SendSLICCmd(string slicCmd)
        {
            try
            {
                mFIC.WriteString(slicCmd, true);
                return true;
            }
            catch (Exception oExp)
            {
                return false;
            }
        }


        /**************************************************************************
         Send the Query SCPI command to the Instrument and reads the Response
         This function returns the string recieved from Instrument's Output Buffer if 
         successful else exception message is returned.
        **************************************************************************/

        public string QueryString(string slicCmd)
        {
            try
            {
                string str;

                // To send SCPI to the Instrument
                SendSLICCmd(slicCmd);

                // To Read output buffer from the Instrument
                str = ReadResult();

                return str;
            }
            catch (Exception oExp)
            {
                // In case of any exception, return string contains exception message
                return (oExp.Message);
            }
        }


        /************************************************************
         This Function Reads output buffer from the Instrument. 
         If any exception is raised, Exception message is returned.
         ************************************************************/

        public string ReadResult()
        {
            try
            {
                string str = mFIC.ReadString();

                // Remove lf character
                str = str.Substring(0, str.Length - 1);

                return str;
            }
            catch (Exception oExp)
            {
                // In case of any exception, return string contains exception message
                return (oExp.Message);
            }
        }


        /********************************************************************
         Sets the VISA Timeout. 
         As few measurement takes time in execution. 
         Thus, VISA Interface has to wait until execution is not completed. 
         ********************************************************************/

        public int TimeOut
        {
            // Returns the Current VISA Timeout
            get { return mFIC.IO.Timeout; }

            // Set the VISA Timeout to the <value> passed
            set { mFIC.IO.Timeout = value; }
        }

    }
}

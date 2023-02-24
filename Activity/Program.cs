using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;

//include System.drawing as a refrence

namespace MyActivity
{
    class ActivityAction
    {
        private static int mouseMovementDelay = 5 * 1000;   //5 seconds
        private static int InitialDelay = 5 * 1000;   //5 seconds
        private static Thread JigThread;
        private static bool doJiggle = true;
        private static bool goUp = false;
        private static int movementDistance = 1;
        private static string sendKey = "^";
        private static POINT previouslpPoint;

        static void Main(string[] args)
        {
            //step one, see if user inputed their own stuff
            processArgs(args);

            System.Diagnostics.Debug.WriteLine("static main");
            ActivityAction myMouse = new ActivityAction();
            while (myMouse.NoUserInteractionDetected2())
            {
                System.Diagnostics.Debug.WriteLine("inside while loop");
            }
            System.Diagnostics.Debug.WriteLine("end");
            doJiggle = false;
            Environment.Exit(0);
        }

        private static void processArgs(string[] args)
        {
            //If this program is run from the commandline, you can feed it arguments
            //currently I'm allowing the following
            //InitDelay
            //MovementDelay
            //MovementRange
            //SendKey
            if (args == null || args.Length == 0)
            {
                System.Diagnostics.Debug.WriteLine("Default params will be used.");
                System.Diagnostics.Debug.WriteLine("To see options pass in the single parameter 'help' or '?'");
            }
            else
            {
                if ((args[0].ToLower().Trim() == "help") || (args[0].Contains("?")))
                {
                    string msg = "This application accepts three arguments, and a help request\n" +
                            "You can pass InitDelay [int] MovementDelay [int] MovementRange [int] SendKey [char] || {special}\n" +
                            "For more description per parameter pass a 'help' or a '?' as the subsequent parameter.\n" +
                            "EG pass 'InitDelay ?' to view InitDelay description and examples.\n\n" +
                            "An example to invoke this program would be\n" +
                            "'InitDelay 15 MovementDelay 5 MovementRange 5'\n" +
                            "Where initial delay shall be 15 seconds, all mouse movement delays will be 5 seconds, and mouse shall move 5 pixels";
                    System.Diagnostics.Debug.WriteLine(msg);
                }
                for (int i = 0; i < args.Length; i += 2)
                {
                    string cmd = args[i].ToLower();
                    switch (cmd)
                    {
                        case "initdelay":
                            InitDelaySet(args[i + 1]);
                            break;
                        case "movementdelay":
                            MovementDelaySet(args[i + 1]);
                            break;
                        case "movementrange":
                            MovementRangeSet(args[i + 1]);
                            break;
                        case "sendkey":
                            SendKeySet(args[i + 1]);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private static void InitDelaySet(string value)
        {
            if (isHelp(value))
            {
                string msg = "Initial delay accepts a positive whole number.\n" +
                    "That number shall be interpreted as seconds. Default value is 5 seconds";
                System.Diagnostics.Debug.WriteLine(msg);
            }
            else
            {
                if (int.TryParse(value, out int result))
                {
                    if (result <= 0)
                    {
                        string msg = $"The parsed value of [{result}] is a negative or zero. " +
                            $"Will default initial time delay to [{InitialDelay}]";
                        System.Diagnostics.Debug.WriteLine(msg);
                    }
                    else
                    {
                        string msg = $"Will set initial time delay to [{result}]";
                        System.Diagnostics.Debug.WriteLine(msg);
                        InitialDelay = result * 1000;
                    }
                }
                else
                {
                    string msg = $"The string [{value}] could not be parsed as an int";
                    System.Diagnostics.Debug.WriteLine(msg);
                }
            }
        }

        private static void MovementDelaySet(string value)
        {
            if (isHelp(value))
            {
                string msg = "Movement delay accepts a positive whole number.\n" +
                    "That number shall be interpreted as seconds. Default value is 5 seconds";
                System.Diagnostics.Debug.WriteLine(msg);
            }
            else
            {
                if (int.TryParse(value, out int result))
                {
                    if (result <= 0)
                    {
                        string msg = $"The parsed value of [{result}] is a negative or zero. " +
                            $"Will default initial delay to [{mouseMovementDelay}]";
                        System.Diagnostics.Debug.WriteLine(msg);
                    }
                    else
                    {
                        string msg = $"Will set movement time delay to [{result}]";
                        System.Diagnostics.Debug.WriteLine(msg);
                        mouseMovementDelay = result * 1000;
                    }
                }
                else
                {
                    string msg = $"The string [{value}] could not be parsed as an int";
                    System.Diagnostics.Debug.WriteLine(msg);
                }
            }
        }

        private static void MovementRangeSet(string value)
        {
            if (isHelp(value))
            {
                string msg = "Movement Range accepts a positive whole number.\n" +
                    "That number shall be interpreted as pixels. Default value is 1 (pixels)";
                System.Diagnostics.Debug.WriteLine(msg);
            }
            else
            {
                if (int.TryParse(value, out int result))
                {
                    if (result <= 0)
                    {
                        string msg = $"The parsed value of [{result}] is a negative or zero. " +
                            $"Will default movment range to [{movementDistance}]";
                        System.Diagnostics.Debug.WriteLine(msg);
                    }
                    else
                    {
                        string msg = $"Will set movement range to [{result}] pixels";
                        System.Diagnostics.Debug.WriteLine(msg);
                        movementDistance = result;
                    }
                }
                else
                {
                    string msg = $"The string [{value}] could not be parsed as an int";
                    System.Diagnostics.Debug.WriteLine(msg);
                }
            }
        }

        private static void SendKeySet(string value)
        {
            if (isHelp(value))
            {
                string msg = "The SendKey command accepts an string that you passed.\n" +
                    "The program does accept special characters, which will need to be encapsalated with {},\n" +
                    "An example would be the ctrl key, you would need to pass {ctrl}, and program will send that keystoke.\n" +
                    "The special characters are as follows:\n" +
                    "ctrl, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12\n" +
                    "See this page for more info--> https://learn.microsoft.com/en-us/office/vba/Language/Reference/user-interface-help/sendkeys-statement";
                System.Diagnostics.Debug.WriteLine(msg);
            }
            else
            {
                string temp = value.Trim();
                if (temp.StartsWith("{") && temp.EndsWith("}"))
                {
                    //encapsalation example, do the break down thing!
                }
                else
                {
                    if (temp.Length > 0)
                        sendKey = temp;
                }

            }
        }

        private static bool isHelp(string v)
        {
            bool returningBool = false;
            if (v.ToLower() == "help")
                returningBool = true;
            else if (v == "?")
                returningBool = true;
            return returningBool;
        }

        public ActivityAction()
        {
            System.Diagnostics.Debug.WriteLine("mouse constructor");
            //main call
            //at this point you shall wait the specific delayed mouse move time
            //then start moving stuff at the described amount of time

            Task<bool> InitialDelayTask = Task.Run(async delegate
            {
                await Task.Delay(InitialDelay);
                return true;
            });
            InitialDelayTask.Wait();
            //POINT lpPoint = GetCursorPos();
            //SetCursorPos(lpPoint);
            FigureOutIfMovingUpOrDown();
            System.Diagnostics.Debug.WriteLine("post initial delay");
            previouslpPoint = GetCursorPos();
            JigThread = new Thread(DoTheJiggle);
            JigThread.Start();
            System.Diagnostics.Debug.WriteLine("started jiggle thread");
        }

        private void FigureOutIfMovingUpOrDown()
        {
            //overkill defensive QC (Quality control)
            //if mouse position is at or below 100, set goUp to true
            //otherwise set goUp to false
            //this is to prevent the mouse from going outside of the screen scope in edge cases
            //like you leave the mouse at the top, or bottom of the screen after initial delay...

            //FYI goUp is set to false at start....
            POINT lpPoint = GetCursorPos();
            if (lpPoint.Y > (movementDistance))
                goUp = true;
        }

        private void DoTheJiggle()
        {
            while (doJiggle)
            {
                POINT lpPoint = GetCursorPos();

                System.Diagnostics.Debug.WriteLine($"Current location {lpPoint.X},{lpPoint.Y}");

                if (goUp)
                {
                    goUp = false;
                    lpPoint.Y += movementDistance;
                }
                else
                {
                    goUp = true;
                    lpPoint.Y += -movementDistance;
                }

                System.Diagnostics.Debug.WriteLine("moving");
                if (string.IsNullOrEmpty(sendKey) == false)
                    SendKeys.SendWait(sendKey);

                SetCursorPos(lpPoint);
                //POINT lpPoint2 = GetCursorPos();
                //System.Diagnostics.Debug.WriteLine($"new location {lpPoint2.X},{lpPoint2.Y}");
                Task<bool> MoseMoveDelayTask = Task.Run(async delegate
                {
                    await Task.Delay(mouseMovementDelay);
                    return true;
                });
                MoseMoveDelayTask.Wait();
            }
        }

        private bool NoUserInteractionDetected2()
        {
            System.Diagnostics.Debug.WriteLine("start no user interaction");
            bool returningBool = false;
            //get current mouse position
            POINT lpPoint = GetCursorPos();

            if (NoChangeInPosition(previouslpPoint, lpPoint))
            {
                //if here then mouse hasn't moved
                //wait, and then do another check

                Task<bool> MoseMoveDelayTask = Task.Run(async delegate
                {
                    await Task.Delay(mouseMovementDelay);
                    return true;
                });
                MoseMoveDelayTask.Wait();

                lpPoint = GetCursorPos();

                if (NoChangeInPosition(previouslpPoint, lpPoint))
                {
                    //if here, then still no change!
                    returningBool = true;
                    //save the current position as previous
                    previouslpPoint = lpPoint;
                }
                else
                {
                    //change detected, stop jiggle
                    doJiggle = false;
                }
            }
            else
            {
                //change detected, stop jiggle
                doJiggle = false;
            }

            return returningBool;
        }

        //private bool NoUserInteractionDetected()
        //{
        //    System.Diagnostics.Debug.WriteLine("start no user interaction");
        //    bool returningBool = false;
        //    //get current mouse position
        //    POINT lpPoint = GetCursorPos();

        //    Task<POINT> ApplicationsTask = Task.Run(async delegate
        //    {
        //        int delay = mouseMovementDelay + (mouseMovementDelay / 2);
        //        await Task.Delay(delay);
        //        return GetCursorPos();
        //    });
        //    ApplicationsTask.Wait();

        //    System.Diagnostics.Debug.WriteLine("post initial delay");
        //    POINT DelayedPoint = ApplicationsTask.Result;

        //    returningBool = NoChangeInPosition(lpPoint, DelayedPoint);

        //    System.Diagnostics.Debug.WriteLine($"NoUserInteractionDetected - ret value [{returningBool}]");
        //    return returningBool;
        //}

        private bool NoChangeInPosition(POINT pt1, POINT pt2)
        {
            //need to check for relative change
            return ((Math.Abs(pt1.X - pt2.X) <= movementDistance) && (Math.Abs(pt1.Y - pt2.Y) <= movementDistance));
        }

        #region C# API wrapper area

        public static POINT GetCursorPos()
        {
            POINT lpPoint;
            GetCursorPos(out lpPoint);
            return lpPoint;

        }
        private static void SetCursorPos(POINT lpPoint)
        {
            //SetCursorPos(in lpPoint.X, in lpPoint.Y);
            System.Windows.Forms.Cursor.Position = new Point(lpPoint.X, lpPoint.Y);
        }

        #endregion C# API wrapper area

        //play around with this later
        #region low level API access
        /// <summary>
        /// Struct representing a point.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public static implicit operator Point(POINT myPoint)
            {
                return new Point(myPoint.X, myPoint.Y);
            }
        }

        /// <summary>
        /// Retrieves the cursor's position, in screen coordinates.
        /// see https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getcursorpos
        /// </summary>
        /// <see>See MSDN documentation for further information.</see>
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        /// <summary>
        /// Retrieves the cursor's position, in screen coordinates.
        /// see https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setcursorpos
        /// </summary>
        /// <see>See MSDN documentation for further information.</see>
        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(in int X, in int Y);

        #endregion #region low level API access
    }
}

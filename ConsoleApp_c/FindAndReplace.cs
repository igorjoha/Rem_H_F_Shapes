using System;

  public class FindAndReplace : IReplacingCallback
    {
        public bool flag = true;
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            if (flag)
            {
                flag = false;
                return ReplaceAction.Replace;
            }
            else
            {
                return ReplaceAction.Stop;
            }
        }
    }
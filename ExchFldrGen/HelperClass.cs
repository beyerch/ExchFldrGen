using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ExchangeUtilities
{
    public static class StringHelperClass
    {

        public static Dictionary<string, string> ParseCMDLineArgs(this string[] ArgumentsArray, string ParameterChar)
        {

            //Create Command Line Dictionary
            Dictionary<string, string> _dict = new Dictionary<string, string>();
            string _Key = "";

            for (int i = 0; i < ArgumentsArray.Length; i++)
            {
                //Create new dictionary entry
                if (ArgumentsArray[i].StartsWith(ParameterChar))
                {
                    if (_Key != "") _dict.Add(_Key.ToUpper(), "");
                    _Key = ArgumentsArray[i].ToString().Substring(1).Trim();
                }
                else
                {
                    //Not a parameter, but a value
                    if (_Key != "")
                    {
                        _dict.Add(_Key.ToUpper(), ArgumentsArray[i].ToString().Trim());
                        _Key = "";
                    }
                }
            }
            //Catch last parameter
            if (_Key != "") _dict.Add(_Key.ToUpper(), "");

            //Return value
            return _dict;
        }


    }
}

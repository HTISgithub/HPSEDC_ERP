using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


public class functions
{

    string tableHtml = "";
    string status = "";
    string color = "yellow";
    string dayId = "0";

    string weekOff = "N";
    string onLeave = "N";

    public string setBoxColor(string status)
    {
        if (status == "PA")
        {
            color = "yellow colorBlack";
        }

        if (status == "AA")
        {
            color = "red foreColorWhite";
        }

        if (status == "RR")
        {
            color = "brown foreColorWhite";
        }
        else if (status == "PP")
        {
            color = "green foreColorWhite";
        }
        else if (status == "LL")
        {
            color = "gray";
        }
        else if (status == "HH")
        {
            color = "blue foreColorWhite";
        }

        return color;
    }
}
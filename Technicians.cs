using System.Collections.Generic;

public class Technicians
{
    public List<string> Info(string t)
    {
        List<string> result = new List<string>() { "null", "null", "null" };
        string nsw = "New South Wales";
        string sa = "South Australia";
        string qld = "Queensland";
        string vic = "Victoria";
        switch (t)
        {
            case "A4":
                result[0] = vic;
                result[1] = "JOHN DEPOILLY";
                result[2] = "DEPOILY JOHN";
                break;
            case "A8":
                result[0] = vic;
                result[1] = "IAN GREENLAND";
                result[2] = "GREENLAND IAN";
                break;
            case "A7":
                result[0] = vic;
                result[1] = "BINDESH PATEL";
                result[2] = "BINDESH PATEL";
                break;
            case "D2":
                result[0] = nsw;
                result[1] = "SCOTT HAY";
                result[2] = "HAY SCOTT";
                break;
            case "D6":
                result[0] = nsw;
                result[1] = "DAVID LEIGH";
                result[2] = "LEIGH DAVID";
                break;
            case "D13":
                result[0] = nsw;
                result[1] = "TECHNICIAN NSW";
                result[2] = "SANTILLO JOE";
                break;
            case "D18":
                result[0] = nsw;
                result[1] = "RITHLESH NARAYAN";
                result[2] = "NARAYAN RITHLESH";
                break;
            case "D19":
                result[0] = nsw;
                result[1] = "ROBERT WICKHAM";
                result[2] = "WICKHAM ROBERT";
                break;
            case "D21":
                result[0] = nsw;
                result[1] = "BIEN SALVADOR";
                result[2] = "SALVADOR BIEN";
                break;

            case "F18":
                result[0] = qld;
                result[1] = "TECHNICIAN QLD";
                result[2] = "PETER FELISE";
                break;
            case "K4":
                result[0] = sa;
                result[1] = "KAM MANI";
                result[2] = "null";
                break;

        }
        return result;

    }
}

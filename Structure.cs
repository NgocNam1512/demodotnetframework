using System;

namespace Reconstructor
{
    public class Location
    {
        public int x1 { get; set; }
        public int y1 { get; set; }
        public int x2 { get; set; }
        public int y2 { get; set; }
    }
    public class ObjData
    {
        public string label { get; set; }
        public Location location { get; set; }
        public string content { get; set; }
    }
    public class Data
    {
        public int width { get; set; }
        public ObjData[] datalist { get; set; }
    }
}

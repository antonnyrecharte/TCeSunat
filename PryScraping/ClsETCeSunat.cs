using System;

namespace PryScraping
{
    internal class ClsETCeSunat
    {
        public ClsETCeSunat(int year = 0, int month = 0)
        {
            TCSanio = year;
            TCSmes = month;
            TCScompra = "0.00";
            TCSventa = "0.00";
        }

        public int TCSanio { get; internal set; }
        public int TCSmes { get; internal set; }
        public int TCSdia { get; internal set; }
        public string TCScompra { get; internal set; }
        public string TCSventa { get; internal set; }
        public string TCSnombreMes { get; internal set; }
        public DateTime TCSfecha { get; internal set; }
    }
}
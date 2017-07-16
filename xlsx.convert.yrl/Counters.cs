using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace xlsx.convert.yrl
{
    public class Counters
    {
        private List<Exception> _exceptions;
        public Counters()
        {
            _exceptions = new List<Exception>();
        }

        public IEnumerable<Exception> Exceptions => _exceptions;

        public void ExceptionInc(Exception ex)
        {
            _exceptions.Add(ex);
        }
    }
}

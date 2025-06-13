using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spravka.Models
{
    public class CertificateResponse
    {
        public bool Success { get; set; }
        public string DocumentUrl { get; set; }
        public string PdfUrl { get; set; }
        public string CertificateNumber { get; set; }
        public string Error { get; set; }
    }

}

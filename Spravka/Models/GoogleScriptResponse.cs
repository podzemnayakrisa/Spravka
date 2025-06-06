using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Spravka.Models
{
    public class GoogleScriptResponse<T>
    {
        public bool Success { get; set; }
        public T Data { get; set; }
    }

}

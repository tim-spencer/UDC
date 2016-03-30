using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UDC.Models;

namespace UDC.Controllers
{
    public class TiffController : ApiController
    {
        // GET: api/Tiff
        public IEnumerable<string> Get()
        {
            return new string[] { "Houston, we have a problem ...", "Toto, this isn't Kansas anymore ..." };
        }

        // GET: api/Tiff?file=<filename>
        public string Get(string file)
        {
            TiffConversion.ConvertToTiff(file);
            return file;
        }
    }
}

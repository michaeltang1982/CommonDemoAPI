using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Sierra.Azure.CommonDemoAPI.Controllers.IDoThis
{
    public partial class IDoThisController : ApiController
    {
        private Models.IDoThis.IDoThisRepository _repository;

        public IDoThisController(Models.IDoThis.IDoThisRepository repository)
        {
            _repository = repository;
        }

        [HttpGet]
        public string Hello()
        {
            return "Hello from I DO THIS";
        }

    }
}

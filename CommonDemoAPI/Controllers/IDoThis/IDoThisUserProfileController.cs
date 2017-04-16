using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Sierra.Azure.CommonDemoAPI.Models.IDoThis;
using System.Threading.Tasks;

namespace Sierra.Azure.CommonDemoAPI.Controllers.IDoThis
{
    public partial class IDoThisController
    {
        [HttpGet]
        [Route("api/IDoThis/UserProfile")]
        public async Task<UserProfile> GetUserProfile(string id)
        {
            return await _repository.GetUserProfile(id);
        }

        [HttpPost]
        [Route("api/IDoThis/UserProfile/Create")]
        public void PostUserProfile([FromBody]string value)
        {
        }

        [HttpPut]
        [Route("api/IDoThis/UserProfile/Update")]        
        public void PutUserProfile(int id, [FromBody]string value)
        {
        }

        [HttpDelete]
        [Route("api/IDoThis/UserProfile/Delete")]
        public void DeleteUserProfile(int id)
        {
        }
    }
}

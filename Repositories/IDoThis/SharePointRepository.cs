using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sierra.Azure.CommonDemoAPI.Models.IDoThis;

namespace Sierra.Azure.CommonDemoAPI.Repositories.IDoThis
{
    public class SharePointRepository : IDoThisRepository
    {
        public UserProfile GetUserProfile(string userId)
        {
            return new UserProfile { Id = userId, Name = "user from SharePoint repository" };
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sierra.Azure.CommonDemoAPI.Models.IDoThis
{
    public interface IDoThisRepository
    {
        Task<UserProfile> GetUserProfile(string userId);
    }
}

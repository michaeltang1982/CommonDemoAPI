using System.Web.Http;
using Microsoft.Practices.Unity;
using Unity.WebApi;
using System.Configuration;

namespace Sierra.Azure.CommonDemoAPI
{
    public static class UnityConfig
    {
        public static void RegisterComponents()
        {
			var container = new UnityContainer();

            // register all your components with the container here
            // it is NOT necessary to register your controllers

            // e.g. container.RegisterType<ITestService, TestService>();
            string configString = ConfigurationManager.AppSettings["IDoThisRepositoryConfig"];
            Repositories.IDoThis.DatabaseRepository repository = new Repositories.IDoThis.DatabaseRepository(configString);
            container.RegisterInstance<Models.IDoThis.IDoThisRepository>(repository);

            

            GlobalConfiguration.Configuration.DependencyResolver = new UnityDependencyResolver(container);
        }
    }
}
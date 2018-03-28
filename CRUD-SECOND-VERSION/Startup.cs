using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CRUD_SECOND_VERSION.Startup))]
namespace CRUD_SECOND_VERSION
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}

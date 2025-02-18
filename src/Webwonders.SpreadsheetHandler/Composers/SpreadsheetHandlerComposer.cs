using Microsoft.Extensions.DependencyInjection;
using Umbraco.Cms.Core.Composing;
using Umbraco.Cms.Core.DependencyInjection;

namespace Webwonders.SpreadsheetHandler;


public class SpreadsheetHandlerComposer : IComposer
{
    public void Compose(IUmbracoBuilder builder)
    {
        builder.Services.AddScoped<IWebwondersSpreadsheetHandler, WebwondersSpreadsheetHandler>();
    }
}

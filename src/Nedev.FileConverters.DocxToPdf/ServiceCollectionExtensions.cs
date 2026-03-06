using Microsoft.Extensions.DependencyInjection;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.DocxToPdf
{
    /// <summary>
    /// Helpers for registering the DOCX→PDF converter in an IoC container.
    /// </summary>
    public static class ServiceCollectionExtensions
    {
        /// <summary>
        /// Adds the built-in DOCX→PDF converter to the service collection.
        /// This is a convenience wrapper around <see cref="Nedev.FileConverters.Core.ServiceCollectionExtensions.AddFileConverter"/>.
        /// </summary>
        /// <param name="services">The service collection.</param>
        /// <returns>The same instance for chaining.</returns>
        public static IServiceCollection AddDocxToPdf(this IServiceCollection services)
        {
            // register an instance, the core infrastructure also supports scanning via attributes
            services.AddFileConverter("docx", "pdf", new DocxToPdfConverter());
            return services;
        }
    }
}
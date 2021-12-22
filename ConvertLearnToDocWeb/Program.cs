using System;
using System.Globalization;
using System.Reflection;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;

namespace ConvertLearnToDocWeb
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                }).Build().Run();
        }

		public static DateTime GetLinkerTime(Assembly assembly = null)
		{
			if (assembly == null)
				assembly = Assembly.GetExecutingAssembly();

			const string BuildVersionMetadataPrefix = "+build";

			var attribute = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
			if (attribute?.InformationalVersion != null)
			{
				var value = attribute.InformationalVersion;
				var index = value.IndexOf(BuildVersionMetadataPrefix);
				if (index > 0)
				{
					value = value[(index + BuildVersionMetadataPrefix.Length)..];
					return DateTime.ParseExact(value, "yyyy-MM-ddTHH:mm:ss:fffZ", CultureInfo.InvariantCulture);
				}
			}
			return default;
		}
	}
}

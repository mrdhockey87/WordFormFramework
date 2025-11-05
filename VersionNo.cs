using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WordFormFramework
{
    public class VersionNo
    {
        public VersionNo()
        {

        }
        public static string GetFrameworkVersion()
        {
            // This will get the assembly containing this class
            var assembly = Assembly.GetExecutingAssembly();

            // Try to get the informational version first (most detailed)
            var infoVersionAttr = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            if (infoVersionAttr != null)
                return infoVersionAttr.InformationalVersion;

            // Fallback to AssemblyVersion
            var version = assembly.GetName().Version?.ToString();

            return version ?? "2.0.5";
        }
    }
}

}
/*Version history:
 *
 *2.0.5 - Fixed version retrieval to use AssemblyInformationalVersion if available. Also fixed save as issue in the rich text editor so 
 *        images are saved correctly without crashing.
 * 1.0.0 - Initial release got framework to open word docx files and convert then to rich text format for editing.
 */
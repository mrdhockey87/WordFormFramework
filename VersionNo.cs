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

            return version ?? "3.2.9";
        }
    }
}
/*Version history:
 *
 * 4.3.12 - Removed the table & protected text support from the rich text editor as it was causing too many issues where the text would 
 *          not show up. The save as rtf works, however it loses the image when saving. The save as Docx just says save failed. mdail 11-6-25
 * 3.2.9 - Added support for open & save rtf docs, add the ability to add protected text that cannot be edited by the user. Added tow
 *         ways to open rtf's with protected data one that allows for locking/unlocking the other does not allow lock text to be edited. mdail 24-4-15
 * 3.1.8 - Added the ability to have tables with fixed column widths in the rich text editor. Also to merge unmerge table cells. 
 *         Also added the ability to open and save rtf docs. mdail 11-6-25
 * 2.0.5 - Fixed version retrieval to use AssemblyInformationalVersion if available. Also fixed save as issue in the rich text editor so 
 *        images are saved correctly without crashing.
 * 1.0.0 - Initial release got framework to open word docx files and convert then to rich text format for editing.
 */
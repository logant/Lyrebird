using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using LMNA.Lyrebird.LyrebirdCommon;

namespace LMNA.Lyrebird
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create the client connection
            LyrebirdChannel channel = new LyrebirdChannel(1);
            channel.Create();
            string message = channel.DocumentName();
            channel.Dispose();
            channel = null;

            Console.WriteLine("Finished: " + message);
        }
    }
}

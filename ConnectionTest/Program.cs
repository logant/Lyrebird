using System;
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

            Console.WriteLine("Finished: " + message);
        }
    }
}

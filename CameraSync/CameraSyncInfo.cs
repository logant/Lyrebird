﻿using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace Lyrebird
{
    public class CameraSyncInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "LyrebirdCameraSync";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                //Return a 24x24 pixel bitmap to represent this GHA library.
                return null;
            }
        }
        public override string Description
        {
            get
            {
                //Return a short string describing the purpose of this GHA library.
                return "";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("066d096b-81b4-4048-8f45-5329c2289d6a");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "Timothy Logan";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "www.github.com/logant";
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;

namespace LMNA.Lyrebird.LyrebirdCommon
{
    public class LyrebirdChannel : IDisposable
    {
        ChannelFactory<ILyrebirdService> _factory;
        ILyrebirdService _channel;
        NetNamedPipeBinding _binding;
        EndpointAddress _endpoint;
        readonly object _locker;
        bool _disposed;
        readonly int productId;

        // Public Constructor
        public LyrebirdChannel(int prodId)
        {
            productId = prodId;
            _locker = new object();
            _disposed = false;
        }

        // Public Creator
        public bool Create()
        {
            bool rc = false;
            try
            {
                _binding = new NetNamedPipeBinding();
                if (productId == 0)
                {
                    //_endpoint = new EndpointAddress("net.pipe://localhost/LMNts/LyrebirdServer/Revit2013/LyrebirdService");
                }
                else if (productId == 2)
                {
                    //_endpoint = new EndpointAddress("net.pipe://localhost/LMNts/LyrebirdServer/Revit2015/LyrebirdService");
                }
                else
                {
                    _endpoint = new EndpointAddress("net.pipe://localhost/LMNts/LyrebirdServer/Revit2014/LyrebirdService");
                }
                _factory = new ChannelFactory<ILyrebirdService>(_binding, _endpoint);
                _channel = _factory.CreateChannel();
                rc = true;
            }
            catch (Exception ex)
            { 
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return rc;
        }

        public string DocumentName()
        {
            if (IsValid())
            {
                try
                {
                    string result = _channel.GetDocumentName();
                    return result;
                }
                catch (Exception ex)
                {
                    Dispose();
                    return ex.Message;
                    //HandleException(ex);
                    //Dispose();
                }
            }
            return null;
        }

        public List<RevitObject> FamilyNames()
        {
            if (IsValid())
            {
                try
                {
                    List<RevitObject> famNames = _channel.GetFamilyNames();
                    return famNames;
                }
                catch
                {
                    List<RevitObject> errors = new List<RevitObject>();
                    errors.Add(new RevitObject("Error", -1, "Error"));
                    return errors;
                }
            }
            return null;
        }

        public List<string> TypeNames(RevitObject familyName)
        {
            if (IsValid())
            {
                try
                {
                    List<string> typeNames = _channel.GetTypeNames(familyName);
                    return typeNames;
                }
                catch
                {
                    List<string> errors = new List<string>();
                    errors.Add("Error");
                    return errors;
                }
            }
            return null;
        }

        public List<RevitParameter> Parameters(RevitObject familyName, string typeName)
        {
            if (IsValid())
            {
                try
                {
                    List<RevitParameter> parameters = _channel.GetParameters(familyName, typeName);
                    return parameters;
                }
                catch
                {
                    List<RevitParameter> errors = new List<RevitParameter>();
                    RevitParameter rp = new RevitParameter("Error", "Error", "Error", false);
                    errors.Add(rp);
                    return errors;
                }
            }
            return null;
        }

        //public bool CreateObjects(List<RevitObject> objects)
        //{
        //    if (IsValid())
        //    {
        //        try
        //        {
        //            bool created = _channel.CreateObjects(objects);
        //            return created;
        //        }
        //        catch
        //        {
        //            return false;
        //        }
        //    }
        //    return false;
        //}

        public bool CreateOrModify(List<RevitObject> objects, Guid uniqueId)
        {
            if (IsValid())
            {
                try
                {
                    bool finished = _channel.CreateOrModify(objects, uniqueId);
                    return finished;
                }
                catch
                {
                    return false;
                }
            }
            return false;
        }

        public bool IsValid()
        {
            if (null != _factory && null != _channel && false == _disposed)
                return true;
            return false;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    lock (_locker)
                    {
                        if (null != _channel)
                        {
                            ((IClientChannel)_channel).Abort();
                            _channel = null;
                        }

                        if (null != _factory)
                        {
                            _factory.Abort();
                            _factory = null;
                        }

                        _endpoint = null;
                        _binding = null;
                    }
                    _disposed = true;
                }
            }
        }
    }
}

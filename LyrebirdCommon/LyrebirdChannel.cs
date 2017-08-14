using System;
using System.Collections.Generic;
using System.ServiceModel;

namespace Lyrebird
{
    public class LBChannel : IDisposable
    {
        ChannelFactory<ILyrebirdService> _factory;
        ILyrebirdService _service;
        NetNamedPipeBinding _binding;
        EndpointAddress _endPoint = null;
        readonly object _locker;
        bool _disposed;
        readonly string _productId;

        public LBChannel(string prodId)
        {
            _productId = prodId;
            _locker = new object();
            _disposed = false;
        }

        /// <summary>
        /// Create the new endpoint for the service and make sure the ChannelFactory and ILyrebirdService are set.
        /// The LBChannel will be used by the Client to connect to the Server which actually implements the ILyrebirdService.
        /// </summary>
        /// <returns>Successfully created?</returns>
        public bool Create()
        {
            bool created = false;
            try
            {
                _binding = new NetNamedPipeBinding();
                _binding.MaxReceivedMessageSize = 52428800;
                _endPoint = new EndpointAddress(Properties.Settings.Default.BaseAddress + "/" + _productId);
                if (null != _endPoint && !_endPoint.IsNone)
                {
                    _factory = new ChannelFactory<ILyrebirdService>(_binding, _endPoint);
                    _service = _factory.CreateChannel();
                    if (_factory.State == CommunicationState.Opened)
                        created = true;
                    else
                        created = false;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error:\n" + ex.Message);
            }

            return created;
        }

        /// <summary>
        /// LBChannel call of the ILyrebirdService GetDocumentName method.
        /// </summary>
        /// <returns></returns>
        public string GetDocumentName()
        {
            if (IsValid())
            {
                try
                {
                    string docName = null;
                    if (_service.GetDocumentName(out docName))
                        return docName;
                    else
                        return "Document Not Found";
                }
                catch (EndpointNotFoundException)
                {
                    Dispose();
                    return "Lyrebird Service for " + _productId + " is not open\nEndpointAddress: " + Properties.Settings.Default.BaseAddress + "/" + _productId;
                }
                catch (Exception ex)
                {
                    Dispose();
                    return ex.Message;
                }
            }
            return "Lyrebird Service is not open.";
        }

        /// <summary>
        /// LBChannel call of the ILyrebirdService GetFamilies Method
        /// </summary>
        /// <returns></returns>
        public List<RevitObject> GetFamilyNames()
        {
            if (IsValid())
            {
                try
                {
                    List<RevitObject> families;
                    if (_service.GetFamilies(out families))
                        return families;
                    else
                        return null;
                }
                catch (EndpointNotFoundException ex)
                {
                    Dispose();
                    return new List<RevitObject> { new RevitObject("Error", -1, "Lyrebird Service is not open.") };
                }
                catch (Exception ex)
                {
                    List<RevitObject> errors = new List<RevitObject> { new RevitObject("Error", -1, ex.Message) };
                    return errors;
                }
            }
            return new List<RevitObject> { new RevitObject("Error", -1, "Lyrebird Service is not open.") };
        }

        /// <summary>
        /// LBChannel call of the ILyrebirdService GetTypeNames method
        /// </summary>
        /// <param name="family"></param>
        /// <returns></returns>
        public List<string> GetTypeNames(RevitObject family)
        {
            if (IsValid())
            {
                try
                {
                    List<string> types;
                    if (_service.GetTypeNames(family.FamilyName, out types))
                        return types;
                    else
                        return null;
                }
                catch (EndpointNotFoundException ex)
                {
                    Dispose();
                    return new List<string> { "Lyrebird Service is not open." };
                }
                catch (Exception ex)
                {
                    return new List<string> { "Error:\n" + ex.Message };
                }
            }
            return new List<string> { "Lyrebird Service is not open." };
        }

        /// <summary>
        /// LBCHannel call of the ILyrebirdService GetParameters method
        /// </summary>
        /// <param name="family"></param>
        /// <param name="typeName"></param>
        /// <returns></returns>
        public List<RevitParameter> GetParameters(RevitObject family, string typeName)
        {
            if (IsValid())
            {
                try
                {
                    List<RevitParameter> parameters = new List<RevitParameter>();
                    if (_service.GetParameters(family.FamilyName, typeName, out parameters))
                        return parameters;
                    else
                        return null; ;
                }
                catch (EndpointNotFoundException ex)
                {
                    Dispose();
                    return new List<RevitParameter> { new RevitParameter("Error", "Error", "Lyrebird Service is not open", true) };
                }
                catch (Exception ex)
                {
                    return new List<RevitParameter> { new RevitParameter("Error", "Error", ex.Message, true) };
                }
            }
            return new List<RevitParameter> { new RevitParameter("Error", "Error", "Lyrebird Service is not open", true) };
        }


        /// <summary>
        /// LBChannel call of the ILyrebirdService GetCategoryElements method
        /// </summary>
        /// <param name="eic"></param>
        /// <returns></returns>
        public List<string> GetCategoryElements(ElementIdCategory eic)
        {
            if (IsValid())
            {
                try
                {
                    List<string> elements = new List<string>();
                    if (_service.GetCategoryElements(eic, out elements))
                        return elements;
                    else
                        return null;
                }
                catch (EndpointNotFoundException ex)
                {
                    Dispose();
                    return new List<string> { "Lyrebird Service is not open." };
                }
                catch (Exception ex)
                {
                    return new List<string> { ex.Message };
                }
            }
            return new List<string> { "Lyrebird Service is not open." };
        }

        /// <summary>
        /// LBChannel call for the ILyrebirdService for creating or modifying Revit elements.
        /// </summary>
        /// <param name="objects"></param>
        /// <param name="uniqueId"></param>
        /// <param name="nickName"></param>
        /// <returns></returns>
        public bool RecieveData(List<RevitObject> objects, Guid uniqueId, string nickName)
        {
            if (IsValid())
            {
                try
                {
                    bool finished = _service.RecieveData(objects, uniqueId, nickName);
                    return finished;
                }
                catch
                {
                    return false;
                }
            }
            return false;
        }

        public Dictionary<string, object> LBAction(Dictionary<string, object> inputs)
        {
            if (!IsValid()) return null;
            try
            {
//                System.Windows.MessageBox.Show(
//                    "At the start of the LyrebirdCommon.LBChannel.LBAction method, so apparently I'm valid so far");
                var outputs = new Dictionary<string, object>();
                _service.LbAction(inputs, out outputs);
                return outputs;
            }
            catch
            {
                return null;
            }
        }

        public List<string> GetRevitAPIPath()
        {
            if (IsValid())
            {
                try
                {
                    List<string> apiPaths;
                    if (_service.GetApiPath(out apiPaths))
                        return apiPaths;

                    return null;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Exception at API Path:\n" + ex.Message);
                    Dispose();
                    return null;
                }
            }
            return null;
        }

        public bool IsValid()
        {
            if (null != _factory && null != _service && false == _disposed)
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
            if (!_disposed && disposing)
            {
                lock (_locker)
                {
                    if (null != _service)
                    {
                        ((IClientChannel)_service).Abort();
                        _service = null;
                    }

                    if (null != _factory)
                    {
                        _factory.Abort();
                        _factory = null;
                    }

                    _endPoint = null;
                    _binding = null;
                }
                _disposed = true;
            }

        }
    }
}

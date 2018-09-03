using System;
using System.Collections.Generic;
using System.ServiceModel;
using CrmServiceHelpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace CRUD
{
    public class CRUD
    {
        private Guid _accountId;
        private OrganizationServiceProxy _serviceProxy;
        private String UserName;
        private ServerConnection sCon;
        private ServerConnection.Configuration config;

        public class rows
        {
            public String Name { get; set; }
            public int Value { get; set; }
        }

        private string getSsl(Boolean b)
        {
            return b ? "https" : "http";
        }

        public OrganizationServiceContext Service
        {
            get { return new OrganizationServiceContext(_serviceProxy); }
        }

        public string getUserName()
        {
            return UserName;
        }

        public string GetOrganization()
        {
            return config.OrganizationName;
        }

        public int Connect(String ServerAddress, String OrganizationName, String Login, String Domain, String Password, Boolean Ssl)
        {
            try
            {
                sCon = new ServerConnection();
                config = sCon.GetServerConfiguration(ServerAddress,
                                                     OrganizationName,
                                                     new Uri(getSsl(Ssl) + "://" + ServerAddress + "/XRMServices/2011/Discovery.svc"),
                                                     new Uri(getSsl(Ssl) + "://" + ServerAddress + "/"+ OrganizationName + "/XRMServices/2011/Organization.svc"),
                                                     "ActiveDirectory",
                                                     Login,
                                                     Domain,
                                                     Password);
                UserName = RunUser(Login, Domain);
                return 1;
            }
            catch (System.ServiceModel.Security.SecurityNegotiationException ex)
            {
                Message.Show("Введен неверный логин/пароль. " + ex.Message, "Ошибка");
                return 2;
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Приложение завершилось с ошибкой: " + ex.Detail.Message, "Ошибка");
                return 3;
            }
            catch (System.TimeoutException ex)
            {
                Message.Show("Превышен лимит ожидания. " + ex.Message, "Ошибка");
                return 4;
            }
            catch (System.Exception ex)
            {
                Message.Show("Приложение завершилось с системной ошибкой: " + ex.Message + " " + ex.InnerException.Message, "Ошибка");
                return 5;
            }
        }

        public void CreateRowEntity(String entityName)
        {
            try
            {
                Entity entity = new Entity(entityName);
                _accountId = _serviceProxy.Create(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Запись не создалась. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void CreateRow(Entity entity)
        {
            try
            {
                _accountId = _serviceProxy.Create(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Запись не создалась. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public Guid getAccount()
        {
            return _accountId;
        }

        public void UpdateRow(String entityName, String column, String value)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, String value, Guid guid)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, guid, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, Decimal value)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, Decimal value, Guid guid)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, guid, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, DateTime value)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, DateTime value, Guid guid)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, guid, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, int value)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                entity[column] = new OptionSetValue(value);
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, Boolean value)
        {
            try
            {
                ColumnSet cols = new ColumnSet(column);
                Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                entity[column] = value;
                _serviceProxy.Update(entity);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        public void UpdateRow(String entityName, String column, String fromEntity, String fromField, String fetch)
        {
            try
            {
                Guid g = RunFetch(fetch, fromField);
                if (g != Guid.Empty)
                {
                    ColumnSet cols = new ColumnSet(column);
                    Entity entity = _serviceProxy.Retrieve(entityName, _accountId, cols);
                    entity[column] = new EntityReference(fromEntity, g);
                    _serviceProxy.Update(entity);
                }
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                Message.Show("Поле '" + column + "' не обновилось. " + ex.Message, "Ошибка");
                throw;
            }
        }

        private String RunUser(String strUser, String strDomain)
        {
            String s = "";
            String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='systemuser'>
                            <attribute name='fullname' />
                            <order attribute='fullname' descending='false' />
                            <filter type='or'>
                            <condition attribute='domainname' operator='eq' value='" + strUser + @"@" + strDomain + @"' />
                            <condition attribute='domainname' operator='eq' value='" + strDomain + @"\" + strUser + @"' />
                            </filter>
                            </entity>
                            </fetch>";
            try
            {
                _serviceProxy = ServerConnection.GetOrganizationProxy(config);
                EntityCollection result = _serviceProxy.RetrieveMultiple(new FetchExpression(fetch));
                foreach (var c in result.Entities)
                {
                    s = c.Attributes["fullname"].ToString();
                }
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>)
            {
            }
            return s;
        }


        public Guid RunFetch(string fetch2, string field2)
        {
            Guid guid = Guid.Empty;
            EntityCollection result = _serviceProxy.RetrieveMultiple(new FetchExpression(fetch2));
            foreach (var c in result.Entities)
            {
                guid = Guid.ParseExact(c.Attributes[field2].ToString(), "D");
            }
            return guid;
        }

        public int getOptionSet(String fetch, String attribute)
        {
            int s = 0;
            try
            {
                _serviceProxy = ServerConnection.GetOrganizationProxy(config);
                EntityCollection result = _serviceProxy.RetrieveMultiple(new FetchExpression(fetch));
                foreach (var c in result.Entities)
                {
                    OptionSetValue CountryOptionSet = c.Attributes.Contains(attribute) ? c[attribute] as OptionSetValue : null;
                    if (CountryOptionSet != null)
                        s = CountryOptionSet.Value;
                }
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>)
            {
            }
            return s;
        }

        public String getOptionSetText(String fetch, String attribute)
        {
            String s = String.Empty;
            try
            {
                _serviceProxy = ServerConnection.GetOrganizationProxy(config);
                EntityCollection result = _serviceProxy.RetrieveMultiple(new FetchExpression(fetch));
                foreach (var c in result.Entities)
                {
                    s = c.FormattedValues[attribute];
                }
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>)
            {
            }
            return s;
        }

        public List<rows> listOptionSet(String entity, String attribute)
        {
            List<rows> Lrow = new List<rows>();
            RetrieveAttributeRequest request = new RetrieveAttributeRequest
            {
                EntityLogicalName = entity,
                LogicalName = attribute,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse response = (RetrieveAttributeResponse)_serviceProxy.Execute(request);
            PicklistAttributeMetadata picklist = (PicklistAttributeMetadata)response.AttributeMetadata;
            foreach (OptionMetadata rec in picklist.OptionSet.Options)
            {
                Lrow.Add(new rows() { Name = rec.Label.LocalizedLabels[0].Label, Value = Int32.Parse(rec.Value.ToString()) });
            }
            return Lrow;
        }

        public void Disconnect()
        {
            sCon = null;
            config = null;
            UserName = null;
            _accountId = Guid.Empty;
            _serviceProxy = null;
        }
    }
}

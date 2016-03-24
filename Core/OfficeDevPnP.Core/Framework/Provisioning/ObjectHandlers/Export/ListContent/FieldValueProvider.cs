using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Field = Microsoft.SharePoint.Client.Field;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.ListContent
{
    public class FieldValueProvider
    {
        #region Internal Classes

        internal class ObjectFieldValueBase
        {
            public Field Field { get; private set; }
            public Web Web { get; private set; }

            public ObjectFieldValueBase(Field field, Web web)
            {
                this.Field = field;
                this.Web = web;
            }            

            public ClientRuntimeContext Context
            {
                get
                {
                    return this.Web.Context;
                }
            }

            public virtual string GetValidatedValue(object value)
            {
                string str = "";
                if (null != value)
                {
                    str = value.ToString();
                }
                return str;
            }

            public virtual object GetFieldValueTyped(string value)
            {
                object valueTyped = value;
                return valueTyped;
            }
        }

        internal class ObjectFieldValueBoolean :
            ObjectFieldValueBase
        {
            public ObjectFieldValueBoolean(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                if (value is bool)
                {
                    str = ((bool)value) ? "TRUE" : "FALSE";
                }
                else
                {
                    str = base.GetValidatedValue(value);
                }
                return str;
            }
        }

        internal class ObjectFieldValueNumber :
            ObjectFieldValueBase
        {
            public ObjectFieldValueNumber(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                double num = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                string str = Convert.ToString(num, CultureInfo.InvariantCulture);
                return str;
            }
        }

        internal class ObjectFieldValueDateTime :
            ObjectFieldValueBase
        {
            public ObjectFieldValueDateTime(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                if (value is DateTime)
                {
                    str = ((DateTime)value).ToString("o", CultureInfo.InvariantCulture);
                }
                else
                {
                    str = base.GetValidatedValue(value);
                }
                return str;
            }
        }

        internal class ObjectFieldValueLookup :
            ObjectFieldValueBase
        {
            private const string SEPARATOR = ";#";
            public ObjectFieldValueLookup(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                FieldLookupValue lookupValue = value as FieldLookupValue;
                if (null != lookupValue)
                {
                    if (0 < lookupValue.LookupId)
                    {
                        str = lookupValue.LookupId.ToString(CultureInfo.InvariantCulture); //lookupValue.LookupId.ToString(CultureInfo.InvariantCulture) + SEPARATOR + lookupValue.LookupValue;
                    }
                }
                else
                {
                    FieldLookupValue[] lookupValues = value as FieldLookupValue[];
                    if (null != lookupValues)
                    {
                        List<string> parts = new List<string>();
                        foreach (FieldLookupValue val in lookupValues)
                        {
                            parts.Add(val.LookupId.ToString(CultureInfo.InvariantCulture));
                            //parts.Add(val.LookupValue.Replace(";", ";;"));
                        }
                        str = string.Join(SEPARATOR, parts.ToArray());
                    }
                    else
                    {
                        str = base.GetValidatedValue(value);
                    }
                }
                return str;
            }

            public override object GetFieldValueTyped(string value)
            {
                object valueTyped = null;
                FieldLookup fieldLookup = this.Field as FieldLookup;
                if (null != fieldLookup)
                {
                    if (fieldLookup.AllowMultipleValues)
                    {
                        List<FieldLookupValue> itemValues = new List<FieldLookupValue>();
                        string[] parts = value.Split(new string[] { SEPARATOR }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string str in parts)
                        {
                            int id;
                            if (int.TryParse(str, out id) && (0 < id))
                            {
                                itemValues.Add(new FieldLookupValue()
                                {
                                    LookupId = id
                                });
                            }
                        }
                        if (0 < itemValues.Count)
                        {
                            valueTyped = itemValues.ToArray();
                        }
                    }
                    else
                    {
                        int id;
                        if (int.TryParse(value, out id) && (0 < id))
                        {
                            valueTyped = new FieldLookupValue()
                            {
                                LookupId = id
                            };
                        }
                    }
                }
                return valueTyped;
            }            
        }

        internal class ObjectFieldValueUser :
           ObjectFieldValueBase
        {
            public ObjectFieldValueUser(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                FieldUserValue userValue = value as FieldUserValue;
                if (null != userValue)
                {
                    str = GetUserLoginById(userValue);
                }
                else
                {
                    FieldUserValue[] userValues = value as FieldUserValue[];
                    if (null != userValues)
                    {
                        List<string> logins = new List<string>();
                        foreach (FieldUserValue val in userValues)
                        {
                            string login = GetUserLoginById(val);
                            logins.Add(login);
                        }
                        str = string.Join(";", logins.ToArray());
                    }
                    else
                    {
                        str = base.GetValidatedValue(value);
                    }
                }
                return str;
            }

            public override object GetFieldValueTyped(string value)
            {
                object userValue = null;
                FieldUser fieldUser = this.Field as FieldUser;
                if (null != fieldUser)
                {
                    string[] logins = null;
                    if (fieldUser.AllowMultipleValues)
                    {
                        logins = value.Split(';');
                    }
                    else
                    {
                        logins = new string[1]
                        {
                            value
                        };
                    }

                    if (fieldUser.AllowMultipleValues)
                    {
                        List<FieldUserValue> values = new List<FieldUserValue>();
                        foreach (string login in logins)
                        {
                            values.Add(FieldUserValue.FromUser(login));
                        }
                        userValue = values.ToArray();
                    }
                    else
                    {
                        userValue = FieldUserValue.FromUser(logins[0]);
                    }
                }
                return userValue;
            }

            private Dictionary<int, string> m_dictUserCache = null;
            private string GetUserLoginById(FieldUserValue userValue)
            {
                string loginName = "";

                if (null == m_dictUserCache)
                {
                    m_dictUserCache = new Dictionary<int, string>();
                }
                string dictValue = "";
                if (m_dictUserCache.TryGetValue(userValue.LookupId, out dictValue))
                {
                    loginName = dictValue;
                }
                else
                {
                    try
                    {
                        var user = this.Web.GetUserById(userValue.LookupId);

                        this.Context.Load(user, u => u.LoginName);
                        this.Context.ExecuteQuery();
                        loginName = user.LoginName;
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, Constants.LOGGING_SOURCE, "Failed to get user by id. User Title: '{0}', User ID:{1}", userValue.LookupValue, userValue.LookupId);
                    }
                    m_dictUserCache.Add(userValue.LookupId, loginName);
                }
                return loginName;
            }
        }

        internal class ObjectFieldValueURL :
           ObjectFieldValueBase
        {
            public ObjectFieldValueURL(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                FieldUrlValue urlValue = value as FieldUrlValue;
                if (null != urlValue)
                {
                    str = string.Format("{0},{1}", urlValue.Url, urlValue.Description);
                }
                else
                {
                    str = base.GetValidatedValue(value);
                }
                return str;
            }

            public override object GetFieldValueTyped(string value)
            {
                var linkValue = new FieldUrlValue();
                var idx = value.IndexOf(',');
                linkValue.Url = (-1 != idx) ? value.Substring(0, idx) : value;
                linkValue.Description = (-1 != idx) ? value.Substring(idx + 1) : value;
                return linkValue;
            }
        }

        internal class ObjectFieldValueChoice :
           ObjectFieldValueBase
        {
            public ObjectFieldValueChoice(Field field, Web web) :
                base(field, web)
            {
            }
        }

        internal class ObjectFieldValueChoiceMulti :
           ObjectFieldValueBase
        {
            public ObjectFieldValueChoiceMulti(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                string[] values = value as string[];
                if (null != values)
                {
                    str = string.Join(";#", values);
                }
                else
                {
                    str = base.GetValidatedValue(value);
                }
                return str;
            }
        }

        internal class ObjectFieldValueContentTypeId :
           ObjectFieldValueBase
        {
            public ObjectFieldValueContentTypeId(Field field, Web web) :
                base(field, web)
            {
            }
        }

        internal class ObjectFieldValueGeolocation :
           ObjectFieldValueBase
        {
            public ObjectFieldValueGeolocation(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                string str = "";
                FieldGeolocationValue geoValue = value as FieldGeolocationValue;
                if (null != geoValue)
                {
                    str = string.Format("{0},{1},{2},{3}", geoValue.Altitude, geoValue.Latitude, geoValue.Longitude, geoValue.Measure);
                }
                return str;
            }

            public override object GetFieldValueTyped(string value)
            {
                object itemValue = value;
                var geolocationArray = value.Split(',');
                if (geolocationArray.Length == 4)
                {
                    var geolocationValue = new FieldGeolocationValue
                    {
                        Altitude = Double.Parse(geolocationArray[0]),
                        Latitude = Double.Parse(geolocationArray[1]),
                        Longitude = Double.Parse(geolocationArray[2]),
                        Measure = Double.Parse(geolocationArray[3]),
                    };
                    itemValue = geolocationValue;
                }
                return itemValue;
            }
        }

        internal class ObjectFieldValueID:
            ObjectFieldValueBase
        {
            public ObjectFieldValueID(Field field, Web web) :
                base(field, web)
            {
            }

            public override string GetValidatedValue(object value)
            {
                double num = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                string str = Convert.ToString(num, CultureInfo.InvariantCulture);
                return str;
            }
        }

        #endregion //Internal Classes

        #region Fields

        private ObjectFieldValueBase _objectFieldValue = null;

        #endregion //Fields

        #region Properties

        public Field Field { get; private set;}
        public Web Web { get; private set; }

        #endregion //Properties

        #region Constructors

        public FieldValueProvider(Field field, Web web)
        {
            this.Field = field;
            this.Web = web;
        }

        #endregion //Constructors

        #region Methods

        public string GetValidatedValue(object value)
        {
            string dbValue = this.ObjectFieldValue.GetValidatedValue(value);
            return dbValue;
        }

        public object GetFieldValueTyped(string value)
        {
            object valueTyped = this.ObjectFieldValue.GetFieldValueTyped(value);
            return valueTyped;
        }

        #endregion //Methods

        #region Implementation

        private ObjectFieldValueBase ObjectFieldValue
        {
            get
            {
                if (null == this._objectFieldValue)
                {
                    this._objectFieldValue = CreateObjectValueTyped();
                }
                return this._objectFieldValue;
            }
        }

        private ObjectFieldValueBase CreateObjectValueTyped()
        {
            if (this.Field.InternalName == "ID")
            {
                return new ObjectFieldValueID(this.Field, this.Web);
            }
            else
            {
                switch (this.Field.FieldTypeKind)
                {
                    case FieldType.Text:
                    case FieldType.Note:
                        return new ObjectFieldValueBase(this.Field, this.Web);
                    case FieldType.Boolean:
                        return new ObjectFieldValueBoolean(this.Field, this.Web);
                    case FieldType.Counter:
                    case FieldType.Number:
                    case FieldType.Currency:
                        return new ObjectFieldValueNumber(this.Field, this.Web);
                    case FieldType.DateTime:
                        return new ObjectFieldValueDateTime(this.Field, this.Web);
                    case FieldType.Lookup:
                        return new ObjectFieldValueLookup(this.Field, this.Web);
                    case FieldType.User:
                        return new ObjectFieldValueUser(this.Field, this.Web);
                    case FieldType.URL:
                        return new ObjectFieldValueURL(this.Field, this.Web);
                    case FieldType.Choice:
                        return new ObjectFieldValueChoice(this.Field, this.Web);
                    case FieldType.MultiChoice:
                        return new ObjectFieldValueChoiceMulti(this.Field, this.Web);
                    case FieldType.ContentTypeId:
                        return new ObjectFieldValueContentTypeId(this.Field, this.Web);
                    case FieldType.Geolocation:
                        return new ObjectFieldValueGeolocation(this.Field, this.Web);
                    default:
                        return new ObjectFieldValueBase(this.Field, this.Web);
                }
            }
        }

        #endregion //Implementation
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;

namespace CAM_Setup_Sheets
{
    // The class derived from DynamicObject.  
    public class DynamicDictionary : DynamicObject, IDictionary<string, object>
    {

        // The inner dictionary.

        public Dictionary<string, object> dictionary = new Dictionary<string, object>();

        // This property returns the number of elements 

        // in the inner dictionary. 

        public int Count
        {
            get
            {
                return dictionary.Count;
            }
        }

        // If you try to get a value of a property  

        // Not defined in the class, this method is called. 

        public override bool TryGetMember(
            GetMemberBinder binder, out object result)
        {

            // Converting the property name to lowercase 

            // so that property names become case-insensitive. 
            string name = binder.Name;

            // If the property name is found in a dictionary, 

            // set the result parameter to the property value and return true. 

            // Otherwise, return false. 
            return dictionary.TryGetValue(name, out result);
        }

        // If you try to set a value of a property that is 

        // not defined in the class, this method is called. 
        public override bool TrySetMember(
            SetMemberBinder binder, object value)
        {

            // Converting the property name to lowercase 

            // so that property names become case-insensitive.
            dictionary[binder.Name] = value;

            // You can always add a value to a dictionary, 

            // so this method always returns true. 
            return true;
        }

        public void Add(string key, object value)
        {
            dictionary.Add(key, value);
        }

        public void Clear()
        {
            dictionary.Clear();
        }

        public bool Contains(string key)
        {
            return dictionary.ContainsKey(key);
        }

        public IDictionaryEnumerator GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        public bool IsFixedSize
        {
            get { return false; }
        }

        public bool IsReadOnly
        {
            get { return false; }
        }

        public ICollection Keys
        {
            get { return dictionary.Keys; }
        }

        public void Remove(string key)
        {
            dictionary.Remove(key);
        }

        public ICollection Values
        {
            get { return dictionary.Values; }
        }

        public object this[string key]
        {
            get
            {
                if (dictionary.ContainsKey(key))
                    return dictionary[key];
                return null;
            }
            set
            {
                dictionary[key] = value;
            }
        }

        public bool ContainsKey(string key)
        {
            return dictionary.ContainsKey(key);
        }

        ICollection<string> IDictionary<string, object>.Keys
        {
            get { return dictionary.Keys; }
        }

        bool IDictionary<string, object>.Remove(string key)
        {
            return dictionary.Remove(key);
        }

        public bool TryGetValue(string key, out object value)
        {
            value = dictionary[key];
            return true;
        }

        ICollection<object> IDictionary<string, object>.Values
        {
            get { return dictionary.Values; }
        }

        public void Add(KeyValuePair<string, object> item)
        {
            dictionary.Add(item.Key, item.Value);
        }

        public bool Contains(KeyValuePair<string, object> item)
        {
            return dictionary.ContainsKey(item.Key);
        }

        public void CopyTo(KeyValuePair<string, object>[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public bool Remove(KeyValuePair<string, object> item)
        {
            return dictionary.Remove(item.Key);
        }

        IEnumerator<KeyValuePair<string, object>> IEnumerable<KeyValuePair<string, object>>.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return dictionary.GetEnumerator();
        }
    }
}

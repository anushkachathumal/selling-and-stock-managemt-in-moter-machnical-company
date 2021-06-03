using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace SqlServerBackupRestore
{
    public class TestClass
    {
        private static TestClass _TestObject;
        private static Object _locker= new Object();

        private TestClass()
        {
            
        }

        public TestClass GetInstance()
        {
            lock (_locker)
            {
                if (_TestObject == null)
                    _TestObject = new TestClass();
                return _TestObject;
            }
        }
    }
}

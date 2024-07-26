using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.Models
{
    public class Generics
    {
        public object Value { get; set; }

        public Generics(object value)
        {
            Value = value;
        }

        public T GetValue<T>()
        {
            return (T)Value;
        }
    }
}

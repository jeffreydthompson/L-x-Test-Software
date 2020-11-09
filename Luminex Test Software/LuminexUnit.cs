using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Luminex_Test_Software
{

    class CablePin
    {
        string id;
        string wireColor;
        double lowV;
        double highV;
        double measuredV = -100;

        public CablePin(string id, string wireColor, double lowV, double highV)
        {
            this.id = id;
            this.wireColor = wireColor;
            this.lowV = lowV;
            this.highV = highV;
        }

        public double MeasuredVoltage
        {
            get { return this.measuredV; }
            set { this.measuredV = value; }
        }

        public string WireColor
        {
            get { return this.wireColor; }
        }

        public string VoltageRangeLow
        {
            get { return this.lowV.ToString(); }
        }

        public string VoltageRangeHigh
        {
            get { return this.highV.ToString(); }
        }

        public bool passes()
        {
            return (lowV < measuredV) && (measuredV < highV);
        }
    }

    class LuminexUnit
    {
        public Dictionary<string, CablePin> pinDictionary;

        public LuminexUnit()
        {
            pinDictionary = new Dictionary<string, CablePin>
            {
                { "_4140_1", new CablePin("_4140_1", "Orange", 13.5, 16.5) },
                { "_4140_2", new CablePin("_4140_2", "Green", -0.1, 0.1) },
                { "_4140_3", new CablePin("_4140_3", "Green", -0.1, 0.1) },
                { "_4140_4", new CablePin("_4140_4", "Red", 4.9, 5.2) },
                { "_4140_5", new CablePin("_4140_1", "Black", -16.5, -13.5) },

                { "_4130_1", new CablePin("_4130_1", "Black", -0.1, 0.1) },
                { "_4130_2", new CablePin("_4130_2", "Black", -0.1, 0.1) },
                { "_4130_3", new CablePin("_4130_3", "Black", -0.1, 0.1) },
                { "_4130_4", new CablePin("_4130_4", "Black", -0.1, 0.1) },
                { "_4130_5", new CablePin("_4130_5", "Black", -0.1, 0.1) },
                { "_4130_7", new CablePin("_4130_7", "Yellow", 10.8, 13.2) },
                { "_4130_8", new CablePin("_4130_8", "Yellow", 10.8, 13.2) },
                { "_4130_9", new CablePin("_4130_9", "Red", 4.9, 5.2) },
                { "_4130_10", new CablePin("_4130_10", "Red", 4.9, 5.2) },
                { "_4130_11", new CablePin("_4130_11", "Red", 4.9, 5.2) },
                { "_4130_12", new CablePin("_4130_12", "Red", 4.9, 5.2) }
            };
        }

        public void ResetCables()
        {
            foreach(KeyValuePair<string, CablePin> keyValuePair in pinDictionary)
            {
                keyValuePair.Value.MeasuredVoltage = -100;
            }
        }
    }
}

# ExcelToClass
Excel sheet to class model

            var xlsxToModel = new ExcelToClass(@"d:\\smartpol.xlsx","SmartPol");
            
            xlsxToModel.ReadFromExcel();
            
            var list = xlsxToModel.LoadFromDataTable<SmartModel>();
            
            foreach(var i in list)
            {
                Console.WriteLine(i.Name);
                Console.WriteLine(i.Spec45);
            }     

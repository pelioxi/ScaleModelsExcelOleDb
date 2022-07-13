using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ScaleModelsExcel.Model
{
    public class UserInfoModel
    {
        public UserInformation GetUserInformation(string guId)
        {
            ScaleModelsExcelEntities db = new ScaleModelsExcelEntities();
            UserInformation info = (from x in db.UserInformations
                                    where x.GUID == guId
                                    select x).FirstOrDefault();

            return info;
        }

        public void InsertUserInformation(UserInformation info)
        {
            ScaleModelsExcelEntities db = new ScaleModelsExcelEntities();
            db.UserInformations.Add(info);
            db.SaveChanges();
        }
    }
}
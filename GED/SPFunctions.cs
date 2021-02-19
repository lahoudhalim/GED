using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace GED
{
    class SPFunctions
    {

        public static void getStatus (ClientContext ctx , int Id , string listName)
        {
            string envoi = "";
            string status = "";
            string Cycle = "";
           
                List itemList = ctx.Web.Lists.GetByTitle(listName);
                ListItem item = itemList.GetItemById(Id);
                ctx.Load(item);
                ctx.ExecuteQuery();

            status = item["Etat"].ToString();
            envoi = item["Envoyer_x0020_le_x0020_document_x0020_pour_x0020_v_x00e9_rification_x002f_validation"].ToString();
            Cycle = item["Cycle_x0020_de_x0020_vie"].ToString();
            
            switch (status)
            {
                case "Brouillon":
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SendEmail(ctx, (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], "subject", "from", "body", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"]);
                    if(envoi  == "Oui" && Cycle.Contains("Rédaction/Validation"))
                    {
                        item["Etat"] = "En attente de validation";
                        item.Update();

                    }
                    if(envoi == "Oui" && !Cycle.Contains("Rédaction/Validation"))
                    {
                        item["Etat"] = "En attente de vérification";
                        item.Update();
                    }
                    break;
                case "En attente de vérification":
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    break;
                case "En attente de validation":
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    break;
                case "En attente de publication":
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    break;
                case "Publié":
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    break;
                case "En attente de révision":
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"],true);
                    break;


            }
            ctx.ExecuteQuery();

        }

        public static void SPPermission(ClientContext ctx, ListItem item , string role , FieldUserValue[] users, bool createADL)
        {
            foreach(FieldUserValue  user in users)
            {
                User userpermission = ctx.Web.SiteUsers.GetById(user.LookupId);
                item.BreakRoleInheritance(true, true);
                RoleDefinitionBindingCollection collRoleDefinitionBinding = new RoleDefinitionBindingCollection(ctx);
                if (createADL)
                {
                    AddAccusseDeLecture(ctx, item.Id, item.DisplayName, userpermission);
                }
                if (role == "modify")
                {
                    collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Contributor)); //Set permission type

                }
                else if(role == "read")
                {
                    collRoleDefinitionBinding.Add(ctx.Web.RoleDefinitions.GetByType(RoleType.Reader)); //Set permission type
                }
                item.RoleAssignments.Add(userpermission, collRoleDefinitionBinding);
            }
            
        }

        public static void AddAccusseDeLecture (ClientContext ctx , int docID ,  string docName ,User lecteur)
        {
            //ListItem itemToAdd = list.AddItem(itemInfo);
            //itemToAdd["Title"] = activity.Title;
            //title = activity.Title;
            //itemToAdd["TitleENG"] = activity.TitleENG;
            //itemToAdd["DetailActivityENG"] = activity.DescriptionActiviteENG;
            //FieldUserValue[] users = FillUserMultiField(activity.Membres, rowNumber, "Membres", Constants.ACTIVITES_NAME, year, title);
            //if (users.Count() > 0 && users[0] != null)
            //    itemToAdd["TmpMembres"] = users;
            //// itemToAdd["Membres"] = users;
            //// itemToAdd["MFSLA"] = activity.MFSLA;
            //FieldUserValue responsable = FillUserField(activity.Responsable, rowNumber, "Responsable", Constants.ACTIVITES_NAME, year, title);
            //if (responsable != null)
            //    itemToAdd["Responsable"] = responsable;
            //itemToAdd["DetailActivity"] = activity.DescriptionActivite;
            //FieldLookupValue lookupDirection = new FieldLookupValue();
            //if (dicDirections.ContainsKey(activity.Direction))
            //{
            //    lookupDirection.LookupId = dicDirections[activity.Direction];
            //}

            List itemList = ctx.Web.Lists.GetByTitle("");
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem itemtoADD = itemList.AddItem(itemCreateInfo);
            itemtoADD["Title"] = docName;
            itemtoADD["Lecteur"] = lecteur;
           // itemtoADD["UF_x0020_du_x0020_lecteur"] = lecteur[""];
            itemtoADD["Document"] = docID;
            itemtoADD["Document_x0020_lu"] = false;
            itemtoADD["Commentaires_x0020__x0028_inform"] = "";


            itemtoADD.Update();
            ctx.ExecuteQuery();
        }

        public static void SendEmail(ClientContext ctx, FieldUserValue[] users , string Subject, string from, string Body , FieldUserValue[] userReceived)
        {
            List<string> usersEmail = new List<string> { };

            foreach(FieldUserValue user in users)
            {
                if (!userReceived.Contains(user))
                {
                    usersEmail.Add(user.Email);
                }
            }
                try
                {
                    using (ctx)
                    {

                        var emailProperties = new EmailProperties();
                        //Email of authenticated external user
                        emailProperties.To = usersEmail;
                        emailProperties.From = from;
                        emailProperties.Body = Body;
                        emailProperties.Subject = Subject;
                        //emailProperties.CC = cc;
                        Utility.SendEmail(ctx, emailProperties);



                        ctx.ExecuteQuery();



                    }
                }
                catch (Exception ex)
                {



                }
         }
        public static void UpdateReceivedEmail(ClientContext ctx , string list )
        {
            if(list == "Redacteur")
            {

            }
        }
    }
}

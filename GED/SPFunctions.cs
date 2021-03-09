using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Taxonomy;

namespace GED
{
    class SPFunctions
    {

        public static void getStatus (ClientContext ctx , int Id , string listName)
        {
            string envoi = "";
            string status = "";
            string Cycle = "";
            FieldUserValue[] emptyuser = null;
            List<String> CibleIDs = new List<string>();


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
                    // SPPermission(ctx, item, "modify", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SendEmail(ctx, (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], "GED - Un document à été créé", 1, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_"],item);
                    UpdateReceivedEmail(ctx, "Redacteur", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], listName, Id);
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
                    SendEmail(ctx, (FieldUserValue[])item["V_x00e9_rificateurs"], "GED - Demande de vérification d'un document",2, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_"],item);
                    UpdateReceivedEmail(ctx, "Verificateur", (FieldUserValue[])item["V_x00e9_rificateurs"], listName, Id);
                    break;
                case "En attente de validation":
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    SendEmail(ctx, (FieldUserValue[])item["V_x00e9_rificateurs"], "subject",  2, (FieldUserValue[])item["Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_"],item);
                    UpdateReceivedEmail(ctx, "Validateur", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], listName, Id);
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
                    SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], "Subject", 1,(FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_info"],item);
                    UpdateReceivedEmail(ctx, "CibleIndvInfo", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], listName, Id);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    // call async function
                    //SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], "Subject", "from", "body", emptyuser);
                    CibleIDs = GetTaxonomiesId(item, "Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information");
                    foreach(string CibleID in CibleIDs)
                    {
                        Task<string> test =callEmailAzAsync("wf_Get-Emails-from-CH-Pole-UF", CibleID);

                    }
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], "Subject", 1, emptyuser,item);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    SendEmail(ctx, (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], "Subject", 1, (FieldUserValue[])item["Email_x0020_Cible_x0020_indiv_x0020_application"],item);
                    UpdateReceivedEmail(ctx, "CibleIndvApp", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], listName, Id);

                    break;
                case "En attente de révision":
                    SendEmail(ctx, (FieldUserValue[])item["Author"], "Subject", 1, emptyuser,item);
                    SPPermission(ctx, item, "modify", (FieldUserValue[])item["Author"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["R_x00e9_dacteur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["V_x00e9_rificateurs"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Validateur_x0028_s_x0029_"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_information"], false);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_collective_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"], true);
                    SPPermission(ctx, item, "read", (FieldUserValue[])item["Cible_x0028_s_x0029__x0020_individuelle_x0028_s_x0029__x0020_de_x0020_la_x0020_diffusion_x0020_pour_x0020_application"],true);
                    if (item["Passer_x0020_en_x0020_publi_x00e9_"].ToString() == "Yes")
                    {   
                        UpdateStatus(ctx, listName, Id);
                    }
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

            List itemList = ctx.Web.Lists.GetByTitle("Accuss%20de%20lecture");
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

        public static void SendEmail(ClientContext ctx, FieldUserValue[] users , string Subject, int index , FieldUserValue[] userReceived,ListItem item)
        {
            
            
            foreach(FieldUserValue user in users)
            {
                bool notcontain = true; 
                if(userReceived != null)
                {
                    foreach (FieldUserValue recuser in userReceived)
                    {
                        if (user.Email.ToString() == recuser.Email.ToString())
                        {
                            notcontain = false;
                        }
                    }
                }
                
                if (notcontain)
                {
                    List<string> usersEmail = new List<string> { };
                    usersEmail.Add(user.Email.ToString());
                    string body = EmailBody(index, item,user.Email.ToString());
                    try
                    {
                        using (ctx)
                        {

                            var emailProperties = new EmailProperties();
                            //Email of authenticated external user
                            emailProperties.To = usersEmail;
                            emailProperties.From = "process@ghtpdfr.fr";
                            emailProperties.Body = body;
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
            }
         }
        public static void UpdateStatus(ClientContext ctx , string listName , int Id)
        {
            List itemList = ctx.Web.Lists.GetByTitle(listName);
            ListItem item = itemList.GetItemById(Id);       
            item["Etat"] = "Publié";
            item.Update();
            ctx.ExecuteQuery();
        }

        public static void UpdateReceivedEmail(ClientContext ctx, string list, FieldUserValue[] users, string listName, int Id)
        {
            List itemList = ctx.Web.Lists.GetByTitle(listName);
            ListItem item = itemList.GetItemById(Id);
            if (list == "Redacteur")
            {

                item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_R_x00e9_dacteur_x0028_s_x0029_"] = users;


            }
            else if (list == "Verificateur")
            {
                item["Email_x0020_envoy_x00e9__x0020_au_x0028_x_x0029__x0020_V_x00e9_rificateur_x0028_s_x0029_"] = users;

            }
            else if (list == "Validateur")
            {
                item["Email_x0020_envoy_x00e9__x0020_aux_x0020_V_x00e9_rificateur_x0028_s_x0029_"] = users;

            }
            else if (list == "CibleIndvApp")
            {

                item["Email_x0020_Cible_x0020_indiv_x0020_application"] = users;

            }
            else if (list == "CibleIndvInfo")
            {
                item["Email_x0020_Cible_x0020_indiv_x0020_info"] = users;

            }
            item.Update();
            ctx.ExecuteQuery();
        }

        public static string GetAppSetting(string key)
        {
            return Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.Process);
        }


        public static async Task<string> callEmailAzAsync(string wfName, string termGuid)
        {
            string ve = string.Empty;
            string body = "{\"managedmetadataID \": \"" + termGuid + "\"}";
            // string url = "https://prod-08.francecentral.logic.azure.com:443/workflows/9085e566a0c64c6ca3a48811f215d975/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=IOg706o5XUEfo3RIjLQLOaSZi3uqoCKxjqPFc8y6Vwk";
            string url = GetAppSetting("wf_Get-Emails-from-CH-Pole-UF");
            using (var httpClient = new HttpClient())
            {
                var content = new StringContent(body, Encoding.UTF8, "application/json");
                var response = httpClient.PostAsync(url, content).Result;
                string result = response.Content.ReadAsStringAsync().Result;
            }
            return ve;
        }
        //public static string GetTaxonomyId(ListItem item, string fieldName)
        //{

        //    TaxonomyFieldValue taxFieldValue = item[fieldName] as TaxonomyFieldValue;
        //    return taxFieldValue.TermGuid;
        //}
        public static List<string> GetTaxonomiesId(ListItem item, string fieldName)
        {
            List<string> Ids = new List<string>();
            TaxonomyFieldValue[] taxFieldValues = item[fieldName] as TaxonomyFieldValue[];

            foreach (TaxonomyFieldValue taxFieldValue in taxFieldValues)

            {

                Ids.Add(taxFieldValue.TermGuid);

            }
            return Ids;
        }

        public static string EmailBody(int index, ListItem item, string useremail)
        {
            string body = "";
            if (index == 1)
            {
                body = @"Bonjour,
                        Un document à été créé dans la bibliothèque GED: 
                        " +
                          item.ServerRedirectedEmbedUrl + "";

            }
            else if (index == 2)
            {
                body = @"Bonjour,
                        Un document demande à être vérifié dans la bibliothèque GED:
                        " + item.ServerRedirectedEmbedUrl + @"

                        Si vous jugez que le document est vérifié et doit passer en validation, veuillez cliquer sur ce lien:
                        
                        https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/true?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws" +
                        @"
                        
                        Sinon, veuillez cliquer sur ce lien:
                        
                        https://prod-30.francecentral.logic.azure.com/workflows/6ba778559279416580dd5c3cfdef3213/triggers/manual/paths/invoke/" + useremail + "/" + item.Id + "/" + item["Etat"].ToString() + "/false?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=BiOqRLe2hB-pDrBG-hWVk2KMdiD_4wuEE96hiZVEWws";


            }
            else if (index == 3)
            {
                if (index == 1)
                {
                    body = @"Bonjour,
                            Un document a été publié dans la bibliothèque GED:
                        " +
                              item.ServerRedirectedEmbedUrl + "";

                }
            }
            return body;
        }
    }

}

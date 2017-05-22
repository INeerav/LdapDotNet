using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System;
using System.Collections;
using System.Text;
using System.DirectoryServices.AccountManagement;
using System.Configuratio;n
using System.Security.Permissions;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Net;

[assembly: SecurityPermission(SecurityAction.RequestMinimum, Execution = true)]
[assembly: DirectoryServicesPermission(SecurityAction.RequestMinimum)]

namespace ActiveDirectory
{
    public partial class frmAD : Form
    {
        public DirectorySearcher dirSearch = null;

        public frmAD()
        {
            InitializeComponent();
        }

        private void frmAD_Load(object sender, EventArgs e)
        {
            txtUsername.Focus();
            btnSearchUserName.Select();

            txtAddress.Text = GetSystemDomain();            
        }

        private string GetSystemDomain()
        {
            try
            {
                return Domain.GetComputerDomain().ToString().ToLower();
            }
            catch (Exception e)
            {
                e.Message.ToString();
                return string.Empty;
            }
        }
      
        private void GetUserInformation(string username, string passowrd, string domain)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                pnlBlock.BringToFront();
                pnlBlock.Visible = true;

                #region For the searching templates
                using (var context = new PrincipalContext(ContextType.Domain, txtAddress.Text, txtContainer.Text, ContextOptions.Negotiate, txtUsername.Text, txtPassword.Text))
                {
                    Logger objLogger = new Logger("PrincipalContext object has been created");
                    using (UserPrincipal user = new UserPrincipal(context))
                    {
                        user.SamAccountName = txtSearchUser.Text;
                        objLogger.LogWrite("Sam server GetUserInformation: " + txtSearchUser.Text);
                        using (var searcher = new PrincipalSearcher(user))
                        {
                            objLogger.LogWrite("PrincipalSearcher : " + searcher.FindOne().SamAccountName);
                            foreach (var result in searcher.FindAll())
                            {
                                objLogger.LogWrite("GetUserInformation in :");
                                objLogger.LogWrite("in GetUserInformation : " + result.SamAccountName);
                                DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                                objLogger.LogWrite("innn : " + Convert.ToString(de.Properties["samAccountName"].Value));
                                //objLogger.LogWrite("results : ");
                                //pnlBlock.Visible = false;
                                //lblFirstname.Text = Convert.ToString(de.Properties["givenName"].Value);
                                //lblEmailId.Text = Convert.ToString(de.Properties["samAccountName"].Value);

                                PrincipalSearchResult<Principal> groups = result.GetGroups();

                                foreach (Principal item in groups)
                                {
                                    objLogger.LogWrite("Groups: " + item.DisplayName + "   " + item.Name);
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message + " " + ex.InnerException, "Search Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void GetUserInformationforLocal(string username, string passowrd, string domain)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                pnlBlock.BringToFront();
                pnlBlock.Visible = true;

                #region For the searching templates
                //using (var context = new PrincipalContext(ContextType.Domain, txtAddress.Text, txtContainer.Text, ContextOptions.ServerBind, txtUsername.Text, txtPassword.Text))
                using (var context = new PrincipalContext(ContextType.Domain, txtAddress.Text, txtContainer.Text))
                {
                    Logger objLogger = new Logger("PrincipalContext object has been created");
                    using (UserPrincipal user = new UserPrincipal(context))
                    {
                        user.SamAccountName = txtSearchUser.Text;
                        objLogger.LogWrite("SamAccountName search: " + txtSearchUser.Text);
                        using (var searcher = new PrincipalSearcher(user))
                        {
                            objLogger.LogWrite("PrincipalSearcher : " + searcher.FindOne().SamAccountName);
                            foreach (var result in searcher.FindAll())
                            {
                                objLogger.LogWrite("GetUserInformationforLocal in :");
                                objLogger.LogWrite("in GetUserInformationforLocal  : " + result.SamAccountName);
                                DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                                objLogger.LogWrite("innn : " + Convert.ToString(de.Properties["samAccountName"].Value));
                                //objLogger.LogWrite("PrincipalSearcher : " + result.DisplayName);
                                //DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                                //objLogger.LogWrite("results : " + Convert.ToString(de.Properties["givenName"].Value) + Convert.ToString(de.Properties["samAccountName"].Value));
                                //pnlBlock.Visible = false;
                                //lblFirstname.Text = Convert.ToString(de.Properties["givenName"].Value);
                                //lblEmailId.Text = Convert.ToString(de.Properties["samAccountName"].Value);

                                PrincipalSearchResult<Principal> groups = result.GetGroups();

                                foreach (Principal item in groups)
                                {
                                    objLogger.LogWrite("Groups: " + item.DisplayName +"   "+ item.Name);
                                }
                               
                            }
                        }
                    }
                }
                #endregion

               
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " " + ex.InnerException, "Search Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowUserInformation(SearchResult rs)
        {
            Cursor.Current = Cursors.Default;
                        
            pnlBlock.Visible = false;

            if (rs.GetDirectoryEntry().Properties["samaccountname"].Value != null)
                lblUsernameDisplay.Text = "Username : " + rs.GetDirectoryEntry().Properties["samaccountname"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["givenName"].Value != null)
                lblFirstname.Text = "First Name : " +rs.GetDirectoryEntry().Properties["givenName"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["initials"].Value != null)
                lblMiddleName.Text = "Middle Name : " + rs.GetDirectoryEntry().Properties["initials"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["sn"].Value != null)
                lblLastName.Text = "Last Name : " + rs.GetDirectoryEntry().Properties["sn"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["mail"].Value != null)
                lblEmailId.Text = "Email ID : " + rs.GetDirectoryEntry().Properties["mail"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["title"].Value != null)
                lblTitle.Text = "Title : " + rs.GetDirectoryEntry().Properties["title"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["company"].Value != null)
                lblCompany.Text = "Company : " + rs.GetDirectoryEntry().Properties["company"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["l"].Value != null)
                lblCity.Text = "City : " + rs.GetDirectoryEntry().Properties["l"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["st"].Value != null)
                lblState.Text = "State : " + rs.GetDirectoryEntry().Properties["st"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["co"].Value != null)
                lblCountry.Text = "Country : " + rs.GetDirectoryEntry().Properties["co"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["postalCode"].Value != null)
                lblPostal.Text = "Postal Code : " + rs.GetDirectoryEntry().Properties["postalCode"].Value.ToString();

            if (rs.GetDirectoryEntry().Properties["telephoneNumber"].Value != null)
                lblTelephone.Text = "Telephone No. : " + rs.GetDirectoryEntry().Properties["telephoneNumber"].Value.ToString();
            
        }

        private DirectorySearcher GetDirectorySearcher(string username, string passowrd, string domain)
        {           
            if(dirSearch == null)
            {
                try
                {
                    dirSearch = new DirectorySearcher(
                        new DirectoryEntry("LDAP://" + domain, username, passowrd));

                    var credentials = new NetworkCredential(username, passowrd, "LDAP://" + domain);
                    //var ldapidentifier = new LdapDirectoryIdentifier(serverName, port, false, false);
                    //var ldapconn = new LdapConnection(ldapidentifier, credentials);
                }
                catch (DirectoryServicesCOMException e)
                {
                    MessageBox.Show("Connection Creditial is Wrong!!!, please Check.", "Erro Info",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    e.Message.ToString();
                }
                return dirSearch;
            }
            else{
                return dirSearch;
            }
        }

        private SearchResult SearchUserByUserName(DirectorySearcher ds, string username)
        {
            ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(samaccountname=" + username + "))";

            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            SearchResult userObject = ds.FindOne();

            if (userObject != null)
                return userObject;
            else
                return null;         
        }

        private SearchResult SearchUserByEmail(DirectorySearcher ds, string email)
        {
            ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(mail=" + email + "))";

            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            SearchResult userObject = ds.FindOne();

            if (userObject != null)
                return userObject;
            else
                return null;
        }

        private void btnSearchUserName_Click(object sender, EventArgs e)
        {
            Logger objLogger = new Logger("Searching user :" + txtUsername.Text.Trim() + "..." + txtPassword.Text.Trim() + "..." + txtAddress.Text.Trim());
            GetUserInformation(txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtAddress.Text.Trim());               
           
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Logger objLogger = new Logger("local server searching user :" + txtUsername.Text.Trim() + "..." + txtPassword.Text.Trim() + "..." + txtAddress.Text.Trim());
            GetUserInformationforLocal(txtUsername.Text.Trim(), txtPassword.Text.Trim(), txtAddress.Text.Trim());      
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
             string strContainer = string.IsNullOrEmpty(txtContainer.Text)? null:txtContainer.Text;
             ADMethodsAccountManagement ADMethods = new ADMethodsAccountManagement(txtAddress.Text, strContainer, txtUsername.Text, txtPassword.Text);
             bool isvalidate = ADMethods.IsADValidate();
             MessageBox.Show(Convert.ToString(isvalidate), "Search Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
           
        }
    }


    public class ADMethodsAccountManagement
    {
        #region Variables // for the internal purpose

        private string sDomain = string.Empty;
        private string sDefaultOU = null;      
        private string sServiceUser = string.Empty;
        private string sServicePassword = string.Empty;

        #endregion

        public ADMethodsAccountManagement(string Domain, string Container, string UserName, string Password) {
            sDomain = Domain;
            sDefaultOU = Container;
            sServiceUser = UserName;
            sServicePassword = Password;
        }

        #region Validate Methods

        /// <summary>
        /// Validates the username and password of a given user
        /// </summary>
        /// <param name="sUserName">The username to validate</param>
        /// <param name="sPassword">The password of the username to validate</param>
        /// <returns>Returns True of user is valid</returns>
        public bool ValidateCredentials(string sUserName, string sPassword)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();
            return oPrincipalContext.ValidateCredentials(sUserName, sPassword);
        }

        /// <summary>
        /// Checks if the User Account is Expired
        /// </summary>
        /// <param name="sUserName">The username to check</param>
        /// <returns>Returns true if Expired</returns>
        public bool IsUserExpired(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            if (oUserPrincipal.AccountExpirationDate != null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Checks if user exists on AD
        /// </summary>
        /// <param name="sUserName">The username to check</param>
        /// <returns>Returns true if username Exists</returns>
        public bool IsUserExisiting(string sUserName)
        {
            if (GetUser(sUserName) == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Checks if user account is locked
        /// </summary>
        /// <param name="sUserName">The username to check</param>
        /// <returns>Returns true of Account is locked</returns>
        public bool IsAccountLocked(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            return oUserPrincipal.IsAccountLockedOut();
        }
        #endregion

        #region Search Methods

        /// <summary>
        /// Gets a certain user on Active Directory
        /// </summary>
        /// <param name="sUserName">The username to get</param>
        /// <returns>Returns the UserPrincipal Object</returns>
        public UserPrincipal GetUser(string sUserName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();

            UserPrincipal oUserPrincipal =
               UserPrincipal.FindByIdentity(oPrincipalContext, sUserName);
            return oUserPrincipal;
        }

        /// <summary>
        /// Gets a certain group on Active Directory
        /// </summary>
        /// <param name="sGroupName">The group to get</param>
        /// <returns>Returns the GroupPrincipal Object</returns>
        public GroupPrincipal GetGroup(string sGroupName)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext();

            GroupPrincipal oGroupPrincipal =
               GroupPrincipal.FindByIdentity(oPrincipalContext, sGroupName);
            return oGroupPrincipal;
        }

        #endregion

        #region User Account Methods

        /// <summary>
        /// Sets the user password
        /// </summary>
        /// <param name="sUserName">The username to set</param>
        /// <param name="sNewPassword">The new password to use</param>
        /// <param name="sMessage">Any output messages</param>
        public void SetUserPassword(string sUserName, string sNewPassword, out string sMessage)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                oUserPrincipal.SetPassword(sNewPassword);
                sMessage = "";
            }
            catch (Exception ex)
            {
                sMessage = ex.Message;
            }
        }

        /// <summary>
        /// Enables a disabled user account
        /// </summary>
        /// <param name="sUserName">The username to enable</param>
        public void EnableUserAccount(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.Enabled = true;
            oUserPrincipal.Save();
        }

        /// <summary>
        /// Force disabling of a user account
        /// </summary>
        /// <param name="sUserName">The username to disable</param>
        public void DisableUserAccount(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.Enabled = false;
            oUserPrincipal.Save();
        }

        /// <summary>
        /// Force expire password of a user
        /// </summary>
        /// <param name="sUserName">The username to expire the password</param>
        public void ExpireUserPassword(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.ExpirePasswordNow();
            oUserPrincipal.Save();
        }

        /// <summary>
        /// Unlocks a locked user account
        /// </summary>
        /// <param name="sUserName">The username to unlock</param>
        public void UnlockUserAccount(string sUserName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            oUserPrincipal.UnlockAccount();
            oUserPrincipal.Save();
        }

        /// <summary>
        /// Creates a new user on Active Directory
        /// </summary>
        /// <param name="sOU">The OU location you want to save your user</param>
        /// <param name="sUserName">The username of the new user</param>
        /// <param name="sPassword">The password of the new user</param>
        /// <param name="sGivenName">The given name of the new user</param>
        /// <param name="sSurname">The surname of the new user</param>
        /// <returns>returns the UserPrincipal object</returns>
        public UserPrincipal CreateNewUser(string sOU,
           string sUserName, string sPassword, string sGivenName, string sSurname)
        {
            if (!IsUserExisiting(sUserName))
            {
                PrincipalContext oPrincipalContext = GetPrincipalContext(sOU);

                UserPrincipal oUserPrincipal = new UserPrincipal
                   (oPrincipalContext, sUserName, sPassword, true /*Enabled or not*/);

                //User Log on Name
                oUserPrincipal.UserPrincipalName = sUserName;
                oUserPrincipal.GivenName = sGivenName;
                oUserPrincipal.Surname = sSurname;
                oUserPrincipal.Save();

                return oUserPrincipal;
            }
            else
            {
                return GetUser(sUserName);
            }
        }

        /// <summary>
        /// Deletes a user in Active Directory
        /// </summary>
        /// <param name="sUserName">The username you want to delete</param>
        /// <returns>Returns true if successfully deleted</returns>
        public bool DeleteUser(string sUserName)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);

                oUserPrincipal.Delete();
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Group Methods

        /// <summary>
        /// Creates a new group in Active Directory
        /// </summary>
        /// <param name="sOU">The OU location you want to save your new Group</param>
        /// <param name="sGroupName">The name of the new group</param>
        /// <param name="sDescription">The description of the new group</param>
        /// <param name="oGroupScope">The scope of the new group</param>
        /// <param name="bSecurityGroup">True is you want this group 
        /// to be a security group, false if you want this as a distribution group</param>
        /// <returns>Returns the GroupPrincipal object</returns>
        public GroupPrincipal CreateNewGroup(string sOU, string sGroupName,
           string sDescription, GroupScope oGroupScope, bool bSecurityGroup)
        {
            PrincipalContext oPrincipalContext = GetPrincipalContext(sOU);

            GroupPrincipal oGroupPrincipal = new GroupPrincipal(oPrincipalContext, sGroupName);
            oGroupPrincipal.Description = sDescription;
            oGroupPrincipal.GroupScope = oGroupScope;
            oGroupPrincipal.IsSecurityGroup = bSecurityGroup;
            oGroupPrincipal.Save();

            return oGroupPrincipal;
        }

        /// <summary>
        /// Adds the user for a given group
        /// </summary>
        /// <param name="sUserName">The user you want to add to a group</param>
        /// <param name="sGroupName">The group you want the user to be added in</param>
        /// <returns>Returns true if successful</returns>
        public bool AddUserToGroup(string sUserName, string sGroupName)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                GroupPrincipal oGroupPrincipal = GetGroup(sGroupName);
                if (oUserPrincipal == null || oGroupPrincipal == null)
                {
                    if (!IsUserGroupMember(sUserName, sGroupName))
                    {
                        oGroupPrincipal.Members.Add(oUserPrincipal);
                        oGroupPrincipal.Save();
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Removes user from a given group
        /// </summary>
        /// <param name="sUserName">The user you want to remove from a group</param>
        /// <param name="sGroupName">The group you want the user to be removed from</param>
        /// <returns>Returns true if successful</returns>
        public bool RemoveUserFromGroup(string sUserName, string sGroupName)
        {
            try
            {
                UserPrincipal oUserPrincipal = GetUser(sUserName);
                GroupPrincipal oGroupPrincipal = GetGroup(sGroupName);
                if (oUserPrincipal == null || oGroupPrincipal == null)
                {
                    if (IsUserGroupMember(sUserName, sGroupName))
                    {
                        oGroupPrincipal.Members.Remove(oUserPrincipal);
                        oGroupPrincipal.Save();
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if user is a member of a given group
        /// </summary>
        /// <param name="sUserName">The user you want to validate</param>
        /// <param name="sGroupName">The group you want to check the 
        /// membership of the user</param>
        /// <returns>Returns true if user is a group member</returns>
        public bool IsUserGroupMember(string sUserName, string sGroupName)
        {
            UserPrincipal oUserPrincipal = GetUser(sUserName);
            GroupPrincipal oGroupPrincipal = GetGroup(sGroupName);

            if (oUserPrincipal == null || oGroupPrincipal == null)
            {
                return oGroupPrincipal.Members.Contains(oUserPrincipal);
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets a list of the users group memberships
        /// </summary>
        /// <param name="sUserName">The user you want to get the group memberships</param>
        /// <returns>Returns an arraylist of group memberships</returns>
        public ArrayList GetUserGroups(string sUserName)
        {
            ArrayList myItems = new ArrayList();
            UserPrincipal oUserPrincipal = GetUser(sUserName);

            PrincipalSearchResult<Principal> oPrincipalSearchResult = oUserPrincipal.GetGroups();

            foreach (Principal oResult in oPrincipalSearchResult)
            {
                myItems.Add(oResult.Name);
            }
            return myItems;
        }

        /// <summary>
        /// Gets a list of the users authorization groups
        /// </summary>
        /// <param name="sUserName">The user you want to get authorization groups</param>
        /// <returns>Returns an arraylist of group authorization memberships</returns>
        public ArrayList GetUserAuthorizationGroups(string sUserName)
        {
            ArrayList myItems = new ArrayList();
            UserPrincipal oUserPrincipal = GetUser(sUserName);

            PrincipalSearchResult<Principal> oPrincipalSearchResult =
                       oUserPrincipal.GetAuthorizationGroups();

            foreach (Principal oResult in oPrincipalSearchResult)
            {
                myItems.Add(oResult.Name);
            }
            return myItems;
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Gets the base principal context
        /// </summary>
        /// <returns>Returns the PrincipalContext object</returns>
        public PrincipalContext GetPrincipalContext()
        {
            try
            {
                PrincipalContext oPrincipalContext = new PrincipalContext
                   (ContextType.Domain, sDomain, sDefaultOU, ContextOptions.SimpleBind,
                   sServiceUser, sServicePassword);
                return oPrincipalContext;
            }catch(Exception ex){
                PrincipalContext oPrincipalContext = new PrincipalContext
                      (ContextType.Domain, sDomain);
                return oPrincipalContext;
            }
          
        }

        public bool IsADValidate() 
        {
            Logger objlogger = new Logger("Valiadtion has been initiated.....");
            try
            {
                using (var pc = new PrincipalContext(ContextType.Domain, sDomain, sDefaultOU, ContextOptions.Negotiate))
                {
                    objlogger.LogWrite("innn validate");
                    return pc.ValidateCredentials(sServiceUser, sServicePassword, ContextOptions.Negotiate | ContextOptions.SecureSocketLayer);
                }
            }
            catch (Exception ex) {
                MessageBox.Show("Validate?  :" +ex.Message + " " + ex.InnerException, "Validate Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

           
        }

        /// <summary>
        /// Gets the principal context on specified OU
        /// </summary>
        /// <param name="sOU">The OU you want your Principal Context to run on</param>
        /// <returns>Returns the PrincipalContext object</returns>
        public PrincipalContext GetPrincipalContext(string sOU)
        {
            PrincipalContext oPrincipalContext =
               new PrincipalContext(ContextType.Domain, sDomain, sOU,
               ContextOptions.SimpleBind, sServiceUser, sServicePassword);
            return oPrincipalContext;
        }

        #endregion
    }
}

# SynContactoADPortal_Sharepoint.Job

Sharepoint Job C#  [Deployed]
This is a job that connects to Active Directory (LDAP) and search for all users in the company, reads their main attributes (DNI is the Key). Then reads the Human Resource ERP (SAP HCM Web SOAP service) and reads the main atributtes (DNI is the Key). Finally it writes to Shrepoint list the join of this sources resulting in a complete, useful, up-to-date company people directory available in Azure AD, Office 365 and the Sharepoint Intranet.

//****************************************************************************
//
//  File:      B1Db.cs
//
//  Copyright (c) SAP 
//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
//****************************************************************************
using System;
using SAPbobsCOM;

namespace B1WizardBase
{
  /// <summary>
  /// Manages all Metadata actions: UserTables, UserFields, UserKeys and UserDefinedObjects.
  /// </summary>
  /// <remarks>
  /// In your AddOn autogenerated project you will have a class inheriting
  /// form B1Db named projectName_Db.
  /// <para>This class will contain the definition of all the metadata information
  /// you added/removed using the AddOnWizard and will be in charge of creating/ 
  /// removing them every time you run your AddOn.</para>
  /// </remarks>
  /// <example>
  /// This example shows the autogenerated code of the class inheriting from B1Db 
  /// in our addon called MyAddOn.
  /// <para> This class creates two user tables called NS_TAB1 and NS_TAB2, 
  /// one user field called NS_UF inside the B1 table WTR3 and one UDO called NS_WD1</para>
  /// <code lang="Visual Basic">
  /// Public Class MyAddOn_Db
  ///   Inherits B1Db
  ///       
  ///   Public Sub New()
  ///     MyBase.New
  ///     Tables = New B1DbTable() {_
  ///                 New B1DbTable("@NS_TAB1", "My User Table 1", BoUTBTableType.bott_NoObject), 
  ///                 New B1DbTable("@NS_TAB2", "My User Table 2", BoUTBTableType.bott_DocumentLines)}
  ///     Columns = New B1DbColumn() {
  ///                 New B1DbColumn("WTR3", "NS_UF", "NS_UF", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, New B1WizardBase.B1DbValidValue(-1) {}, -1)}
  ///     Udos = New B1Udo() {
  ///                 New B1Udo("NS_WD1", "My UDO", "TTT_WD1", New String() {"TTT_WDL1"}, 
  ///                   BoUDOObjType.boud_Document, BoYesNoEnum.tYES, BoYesNoEnum.tYES, 
  ///                   BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tNO, 
  ///                   BoYesNoEnum.tNO, BoYesNoEnum.tYES, Nothing, 
  ///                   New String() {"DocEntry", "DocNum"}, 
  ///                   New String() {"DocEntry", "DocNum"})}
  ///   End Sub
  /// End Class
  /// </code>
  /// <code lang="C#">
  /// public class MyAddOn_Db : B1Db {
  ///       
  ///   public MyAddOn_Db() {
  ///       Tables = new B1DbTable[] {
  ///               new B1DbTable("@NS_TAB1", "My User Table 1", BoUTBTableType.bott_MasterData),
  ///               new B1DbTable("@NS_TAB2", "My User Table 2", BoUTBTableType.bott_MasterDataLines)};
  ///       Columns = new B1DbColumn[] {
  ///               new B1DbColumn("WTR3", "NS_UF", "NS_UF", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, new B1WizardBase.B1DbValidValue[0], -1)};
  ///       Udos = new B1Udo[] {
  ///               new B1Udo("NS_WD1", "My UDO", "TTT_WD1", new string[] {"TTT_WDL1"}, 
  ///                 BoUDOObjType.boud_Document, BoYesNoEnum.tYES, 
  ///                 BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tYES, 
  ///                 BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, null, 
  ///                 new string[] {"DocEntry", "DocNum"}, new string[] {"DocEntry", "DocNum"})};
  ///   }
  /// }
  /// </code>
  /// </example>
  public class B1Db
  {
    /// <summary>
    /// Collection of UserTables to be added.
    /// </summary>
    protected B1DbTable[]	Tables;
    /// <summary>
    /// Collection of UserFields to be added.
    /// </summary>
    protected B1DbColumn[] Columns;

    /// <summary>
    /// Collection of UserKeys to be added.
    /// </summary>
    protected B1DbKey[] Keys;

    /// <summary>
    /// Collection of UserObjects to be added.
    /// </summary>
    protected B1Udo[]		Udos;


    ////////////////////////////////////////////////////////////////////

    /// <summary>
    /// Default Constructor.
    /// </summary>
    public B1Db()
    {
    }

    /// <summary>
    /// Adds to the company we are connected to all tables, fields and UDOs 
    /// defined by the user using the Wizard.
    /// </summary>
    /// <param name="company">SAPbobsCOM.Company we are connected to (B1Connections.diCompany).</param>
    public void Add(Company company)
    {
      if	(Tables != null)
        foreach (B1DbTable table in Tables)
          table.Add(company);

      if	(Columns != null)
        foreach (B1DbColumn column in Columns)
          column.Add(company);

      if	(Keys != null)
        foreach (B1DbKey key in Keys)
          key.Add(company);

      if	(Udos != null)
        foreach	(B1Udo udo in Udos)
          udo.Add(company);
    }
  }
}
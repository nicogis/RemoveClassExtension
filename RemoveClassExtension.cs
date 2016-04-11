namespace StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension
{
    using System;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using ESRI.ArcGIS.Carto;
    using ESRI.ArcGIS.Catalog;
    using ESRI.ArcGIS.CatalogUI;
    using ESRI.ArcGIS.esriSystem;
    using ESRI.ArcGIS.Geodatabase;

    public class RemoveClassExtension : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        IGxApplication pGxApp = null;

        public RemoveClassExtension()
        {
            this.pGxApp  = ArcCatalog.Application as IGxApplication;
           
        }

        protected override void OnClick()
        {
            try
            {
                if (this.pGxApp.Selection.Count != 1)
                {
                    return;
                }


                IGxObject pGxObject = this.pGxApp.SelectedObject;


                if (!(pGxObject is IGxDataset))
                {
                    return;
                }

                IGxDataset pGxDataset = pGxObject as IGxDataset;
                if (pGxDataset == null)
                {
                    return;
                }

                if (((pGxObject as IGxDataset).Type) != esriDatasetType.esriDTFeatureClass)
                {
                    return;
                }

                try
                {
                    IDataset a = pGxDataset.Dataset;
                }
                catch (COMException COMex)
                {
                    if (COMex.ErrorCode == -2147467259)
                    {
                        if (MessageBox.Show("Not found component register so I don't see the UID: do I have to remove class extension? Are you sure?", "Remove class extension", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            try
                            {
                                IGxObject gxObject = pGxObject.Parent;
                                if (!(gxObject is IGxDatabase2))
                                {
                                    gxObject = gxObject.Parent;
                                }

                                IFeatureWorkspaceSchemaEdit featureWorkspaceSchemaEdit = ((gxObject as IGxDatabase2).Workspace) as IFeatureWorkspaceSchemaEdit;
                                featureWorkspaceSchemaEdit.AlterClassExtensionCLSID(pGxDataset.DatasetName.Name, null, null);

                                IObjectClassDescription featureClassDescription = new FeatureClassDescriptionClass();
                                featureWorkspaceSchemaEdit.AlterInstanceCLSID(pGxDataset.DatasetName.Name, featureClassDescription.InstanceCLSID);

                                MessageBox.Show("Class extension removed: success! Restart ArcCatalog!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch(Exception ex)
                            {
                                MessageBox.Show("Error " + ex.Message);
                                return;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error " + COMex.ErrorCode.ToString() + ": " + COMex.Message);
                    }

                    return;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                    return;
                }


                if (!(pGxDataset.Dataset is IClass))
                {
                    return;
                }

                IClass pClass = pGxDataset.Dataset as IClass;

                if (pClass.EXTCLSID == null)
                {
                    MessageBox.Show("No class extension!");
                    return;
                }
                else
                {

                    IObjectClassDescription ocDescription = new AnnotationFeatureClassDescriptionClass();
                    if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                    {
                        MessageBox.Show("Class extension well-know: I don't remove it (annotation)!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    ocDescription = new DimensionClassDescriptionClass();
                    if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                    {
                        MessageBox.Show("Class extension well-know: I don't remove it (dimension)!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }



                    if (MessageBox.Show(string.Format("Class extension: {0}: do I have to remove class extension?", pClass.EXTCLSID.Value.ToString()), "Remove Class Extension", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (this.RemoveClsExt(pClass))
                        {
                            MessageBox.Show("Class extension removed: success!", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Class extension removed: error! I have problem with exclusive schema lock ", "Remove Class Extension", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString());
            }

        }

        protected override void OnUpdate()
        {
            if (this.pGxApp.Selection.Count != 1)
            {
                this.Enabled = false;
                return;
            }

            IGxObject pGxObject = this.pGxApp.SelectedObject;

            this.Enabled = ((pGxObject as IGxDataset) != null) && ((pGxObject as IGxDataset).Type == esriDatasetType.esriDTFeatureClass);

        }

        private bool FeatureClassWellKnown(UID uid)
        {
            IObjectClassDescription ocDescription = new AnnotationFeatureClassDescriptionClass();
            if ((ocDescription.ClassExtensionCLSID.Value.ToString() == uid.Value.ToString()) && (ocDescription.ClassExtensionCLSID.SubType == uid.SubType))
            {
                return true;
            }

            ocDescription = new DimensionClassDescriptionClass();
            if ((ocDescription.ClassExtensionCLSID.Value.ToString() == uid.Value.ToString()) && (ocDescription.ClassExtensionCLSID.SubType == uid.SubType))
            {
                return true;
            }

            return false;
        }

        private bool RemoveClsExt(IClass classExtension)
        {
            ISchemaLock schemaLock = (ISchemaLock)classExtension;
            bool result = false;
            try
            {
                // Attempt to get an exclusive schema lock.
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);

                // Cast the object class to the IClassSchemaEdit2 interface.
                IClassSchemaEdit2 classSchemaEdit = (IClassSchemaEdit2)classExtension;

                // Clear the class extension.
                classSchemaEdit.AlterClassExtensionCLSID(null, null);

                IObjectClassDescription featureClassDescription = new FeatureClassDescriptionClass();
                classSchemaEdit.AlterInstanceCLSID(featureClassDescription.InstanceCLSID);

                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriSharedSchemaLock);
            }

            return result;
        }
    }

    
}

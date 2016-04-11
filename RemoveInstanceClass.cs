namespace StudioAT.ArcGIS.ArcCatalog.AddIn.RemoveClassExtension
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using ESRI.ArcGIS.CatalogUI;
    using ESRI.ArcGIS.Geodatabase;
    using ESRI.ArcGIS.Catalog;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using ESRI.ArcGIS.esriSystem;
    using ESRI.ArcGIS.Carto;

    public class RemoveInstanceClass : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        IGxApplication pGxApp = null;
        IObjectClassDescription featureClassDescription = null;

        public RemoveInstanceClass()
        {
            this.pGxApp = ArcCatalog.Application as IGxApplication;
            this.featureClassDescription = new FeatureClassDescriptionClass();
            
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
                        if (MessageBox.Show("Not found component register so I don't see the UID: do I have to remove instance class? Are you sure?", "Remove instance class", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            try
                            {
                                IGxObject gxObject = pGxObject.Parent;
                                if (!(gxObject is IGxDatabase2))
                                {
                                    gxObject = gxObject.Parent;
                                }

                                IFeatureWorkspaceSchemaEdit featureWorkspaceSchemaEdit = ((gxObject as IGxDatabase2).Workspace) as IFeatureWorkspaceSchemaEdit;

                                featureWorkspaceSchemaEdit.AlterInstanceCLSID(pGxDataset.DatasetName.Name, this.featureClassDescription.InstanceCLSID);

                                MessageBox.Show("Instance class removed: success! Restart ArcCatalog!", "Remove instance Extension", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
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


                if ((pClass.CLSID.Value.ToString() == this.featureClassDescription.InstanceCLSID.Value.ToString()) && (pClass.CLSID.SubType == this.featureClassDescription.InstanceCLSID.SubType))
                {
                    MessageBox.Show("No instance class found!");
                    return;
                }
                else
                {
                    //if has annotation class extension I set a clsInstance annotation
                    //if has dimension class extension I set a clsInstance dimension
                    IObjectClassDescription ocDescription = new AnnotationFeatureClassDescriptionClass();
                    if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                    {

                        if ((pClass.CLSID.Value.ToString() == ocDescription.InstanceCLSID.Value.ToString()) && (pClass.CLSID.SubType == ocDescription.InstanceCLSID.SubType))
                        {
                            MessageBox.Show("No instance class found!");
                            return;
                        }

                        
                        if (MessageBox.Show("This feature class has extension class Annotation so I update with Annotation instance, confirm?", "Remove instance class", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            if (this.RemoveClsInstance(pClass, ocDescription.InstanceCLSID))
                            {
                                MessageBox.Show("Instance class updated: success!", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Instance class updated: error! I have problem with exclusive schema lock ", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        
                        return;
                    }

                    ocDescription = new DimensionClassDescriptionClass();

                    if ((pClass.EXTCLSID.Value.ToString() == ocDescription.ClassExtensionCLSID.Value.ToString()) && (pClass.EXTCLSID.SubType == ocDescription.ClassExtensionCLSID.SubType))
                    {
                        if ((pClass.CLSID.Value.ToString() == ocDescription.InstanceCLSID.Value.ToString()) && (pClass.CLSID.SubType == ocDescription.InstanceCLSID.SubType))
                        {
                            MessageBox.Show("No instance class found!");
                            return;
                        }
                        
                        if (MessageBox.Show("This feature class has extension class Dimension so I update with Dimension instance, confirm?", "Remove instance class", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            if (this.RemoveClsInstance(pClass, ocDescription.InstanceCLSID))
                            {
                                MessageBox.Show("Instance class updated: success!", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Instance class updated: error! I have problem with exclusive schema lock ", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        return;
                    }
                    
                    
                    if (MessageBox.Show(string.Format("Instance Class: {0}: do I have to remove instance class?", pClass.CLSID.Value.ToString()), "Remove instance class", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (this.RemoveClsInstance(pClass))
                        {
                            MessageBox.Show("Instance class removed: success!", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Instance class removed: error! I have problem with exclusive schema lock ", "Remove instance class", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private bool RemoveClsInstance(IClass classInstance)
        {
            return this.RemoveClsInstance(classInstance, this.featureClassDescription.InstanceCLSID);
        }

        private bool RemoveClsInstance(IClass classInstance, UID uid)
        {
            ISchemaLock schemaLock = (ISchemaLock)classInstance;
            bool result = false;
            try
            {
                // Attempt to get an exclusive schema lock.
                schemaLock.ChangeSchemaLock(esriSchemaLock.esriExclusiveSchemaLock);

                // Cast the object class to the IClassSchemaEdit2 interface.
                IClassSchemaEdit2 classSchemaEdit = (IClassSchemaEdit2)classInstance;
                
                classSchemaEdit.AlterInstanceCLSID(uid);

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

/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpSession;

import org.zkoss.io.Files;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Radiogroup;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.DibujoConfTipologia;
import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.DibujoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;
import com.sener.sireca.web.worker.DrawingWorker;

public class DibujoNewPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button dibujoReplanteo;
    @Wire
    Button volver;
    @Wire
    Button unCheckAll;
    @Wire
    Button checkAll;
    @Wire
    Button uploadFile;
    @Wire
    Textbox fileToUpload;
    @Wire
    Textbox pkInicial;
    @Wire
    Textbox pkFinal;
    @Wire
    Textbox notes;
    @Wire
    Checkbox geoPost;
    @Wire
    Checkbox etiPost;
    @Wire
    Checkbox datPost;
    @Wire
    Checkbox vanos;
    @Wire
    Checkbox flechas;
    @Wire
    Checkbox descentramientos;
    @Wire
    Checkbox implantacion;
    @Wire
    Checkbox altHilo;
    @Wire
    Checkbox distCant;
    @Wire
    Checkbox conexiones;
    @Wire
    Checkbox protecciones;
    @Wire
    Checkbox pendolado;
    @Wire
    Checkbox altCat;
    @Wire
    Checkbox puntSing;
    @Wire
    Checkbox cableado;
    @Wire
    Checkbox datTraz;
    @Wire
    Radiogroup rg;
    @Wire
    Combobox versionList;
    @Wire
    Combobox revisionList;

    private List<ReplanteoVersion> verList;

    private Media media = null;
    private boolean bHDC = true;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);
        fillConf("HDC");

        Project project = projectService.getProjectById(actProj.getIdActive(session));

        verList = replanteoService.getVersions(project);

        List<Integer> vList = replanteoService.getVersionList(project);

        versionList.setModel(new ListModelList(vList));
        versionList.setValue("Escoja Versión");
        revisionList.setValue("Escoja Revisión");

        uploadFile.setDisabled(bHDC);

        uploadFile.setUpload("true");

        uploadFile.addEventListener("onUpload",
                new EventListener<UploadEvent>()
                {
                    @Override
                    public void onEvent(UploadEvent event) throws Exception
                    {
                        media = event.getMedia();
                        String fileName = media.getName();
                        if (fileName.endsWith(".dwg"))
                            fileToUpload.setValue(fileName);
                        else
                            Clients.showNotification("Debe subir un fichero con el formato de la plantilla indicada.");
                    }

                });
    }

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Listen("onChange = #versionList")
    public void fillRevisions()
    {
        revisionList.setValue("Escoja Revisión");
        List<Integer> rList = replanteoService.getRevisionList(verList.get(versionList.getSelectedIndex()));
        revisionList.setModel(new ListModelList(rList));
    }

    @Listen("onClick = #dibujoReplanteo; onOK=#drawingNewWin")
    public void doDraw() throws IOException
    {

        if (!bHDC & fileToUpload.getValue().equals(""))
            Clients.showNotification("No ha seleccionado ningun archivo.");

        else if (pkInicial.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Inicial.");

        else if (pkFinal.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Final.");

        else if (versionList.getValue().equals("Escoja Versión"))
            Clients.showNotification("Debe seleccionar una Version.");

        else if (revisionList.getValue().equals("Escoja Revisión"))
            Clients.showNotification("Debe seleccionar una Revisión.");

        else
        {
            double pkIni = 0;
            double pkFin = 0;

            try
            {
                pkIni = Double.parseDouble(pkInicial.getValue().replace(',',
                        '.'));
                pkFin = Double.parseDouble(pkFinal.getValue().replace(',', '.'));
            }
            catch (Exception e)
            {
                Clients.showNotification("El PK debe ser menor un valor numérico.");
                return;
            }

            if (pkIni >= pkFin)
            {
                Clients.showNotification("El PK Inicial debe ser que el PK Final.");
                return;
            }

            Project project = projectService.getProjectById(actProj.getIdActive(session));
            int numVersion = dibujoService.getLastVersion(project);
            int repVersion = versionList.getSelectedItem().getValue();
            int repRevision = revisionList.getSelectedItem().getValue();
            int idCatenaria = project.getIdCatenaria();

            String catenaria = catenariaService.getCatenariaById(idCatenaria).getNomCatenaria();

            ReplanteoVersion replanteoVersion = replanteoService.getVersion(
                    project, repVersion);
            ReplanteoRevision replanteoRevision = replanteoService.getRevision(
                    replanteoVersion, repRevision);

            DibujoVersion dibujoVersion = dibujoService.getVersion(project,
                    numVersion);
            DibujoRevision dibujoRevision = dibujoService.createRevision(
                    dibujoVersion, replanteoRevision, notes.getValue());

            String ruta = dibujoRevision.getAutoCadPath();

            File dest = new File(ruta);

            if (bHDC)
                Files.copy(
                        dest,
                        new FileInputStream(project.getTemplatePath(DibujoVersion.DIBUJO_REPLANTEO
                                + "HDC.dwg")));

            else
                Files.copy(dest, media.getStreamData());

            DibujoConfTipologia confTipo = new DibujoConfTipologia();

            confTipo.setGeoPost(geoPost.isChecked());
            confTipo.setEtiPost(etiPost.isChecked());
            confTipo.setDatPost(datPost.isChecked());
            confTipo.setVanos(vanos.isChecked());
            confTipo.setFlechas(flechas.isChecked());
            confTipo.setDescentramientos(descentramientos.isChecked());
            confTipo.setImplantacion(implantacion.isChecked());
            confTipo.setAltHilo(altHilo.isChecked());
            confTipo.setDistCant(distCant.isChecked());
            confTipo.setConexiones(conexiones.isChecked());
            confTipo.setProtecciones(protecciones.isChecked());
            confTipo.setPendolado(pendolado.isChecked());
            confTipo.setAltCat(altCat.isChecked());
            confTipo.setPuntSing(puntSing.isChecked());
            confTipo.setCableado(cableado.isChecked());
            confTipo.setDatTraz(datTraz.isChecked());

            DrawingWorker dw = new DrawingWorker(dibujoRevision, confTipo, pkIni, pkFin, bHDC, catenaria);

            dw.start();

            Executions.getCurrent().sendRedirect(
                    "/drawing/progress/" + numVersion + "/"
                            + dibujoRevision.getNumRevision());
        }
    }

    @Listen("onClick = #volver")
    public void doGoBack()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/drawing");
    }

    @Listen("onClick = #checkAll")
    public void doCheckAll()
    {
        geoPost.setChecked(true);
        etiPost.setChecked(true);
        datPost.setChecked(true);
        vanos.setChecked(true);
        flechas.setChecked(true);
        descentramientos.setChecked(true);
        implantacion.setChecked(true);
        altHilo.setChecked(true);
        distCant.setChecked(true);
        conexiones.setChecked(true);
        protecciones.setChecked(true);
        pendolado.setChecked(true);
        altCat.setChecked(true);
        puntSing.setChecked(true);
        cableado.setChecked(true);
        datTraz.setChecked(true);
    }

    @Listen("onClick = #unCheckAll")
    public void doUnCheckAll()
    {
        geoPost.setChecked(false);
        etiPost.setChecked(false);
        datPost.setChecked(false);
        vanos.setChecked(false);
        flechas.setChecked(false);
        descentramientos.setChecked(false);
        implantacion.setChecked(false);
        altHilo.setChecked(false);
        distCant.setChecked(false);
        conexiones.setChecked(false);
        protecciones.setChecked(false);
        pendolado.setChecked(false);
        altCat.setChecked(false);
        puntSing.setChecked(false);
        cableado.setChecked(false);
        datTraz.setChecked(false);
    }

    @Listen("onCheck = #rg")
    public void updateData()
    {
        fillConf(rg.getSelectedItem().getLabel());
    }

    private void fillConf(String type)
    {
        if (type.equals("Replanteo"))
        {
            geoPost.setChecked(true);
            etiPost.setChecked(true);
            datPost.setChecked(false);
            vanos.setChecked(true);
            flechas.setChecked(true);
            descentramientos.setChecked(false);
            implantacion.setChecked(true);
            altHilo.setChecked(false);
            distCant.setChecked(true);
            conexiones.setChecked(false);
            protecciones.setChecked(true);
            pendolado.setChecked(false);
            altCat.setChecked(false);
            puntSing.setChecked(true);
            cableado.setChecked(true);
            datTraz.setChecked(false);
            bHDC = false;
        }
        else if (type.equals("HDC"))
        {
            geoPost.setChecked(true);
            etiPost.setChecked(true);
            datPost.setChecked(false);
            vanos.setChecked(true);
            flechas.setChecked(true);
            descentramientos.setChecked(true);
            implantacion.setChecked(true);
            altHilo.setChecked(true);
            distCant.setChecked(true);
            conexiones.setChecked(true);
            protecciones.setChecked(true);
            pendolado.setChecked(true);
            altCat.setChecked(true);
            puntSing.setChecked(true);
            cableado.setChecked(true);
            datTraz.setChecked(false);
            bHDC = true;
        }

        uploadFile.setDisabled(bHDC);

    }

}

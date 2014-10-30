/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.text.SimpleDateFormat;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Label;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Progressmeter;
import org.zkoss.zul.Timer;

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.DibujoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class DibujoProgressPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Progressmeter postes;
    @Wire
    Label progressLabel;
    @Wire
    Progressmeter function;
    @Wire
    Label funcLabel;
    @Wire
    Label version;
    @Wire
    Label revision;
    @Wire
    Label fecha;
    @Wire
    Timer timer;

    private int numVersion;

    private int numRevision;

    private Project project;

    DibujoVersion ver;
    DibujoRevision rev;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        numVersion = (Integer) Executions.getCurrent().getAttribute(
                "numVersion");
        numRevision = (Integer) Executions.getCurrent().getAttribute(
                "numRevision");

        project = projectService.getProjectById(actProj.getIdActive(session));
        ver = dibujoService.getVersion(project, numVersion);
        rev = dibujoService.getRevision(ver, numRevision);

        version.setValue("Version: " + numVersion);
        revision.setValue("Revision: " + numRevision);
        fecha.setValue("Fecha: "
                + new SimpleDateFormat("dd-MM-yyyy").format(rev.getDate()));

        refreshGrid();

    }

    // TODO: Cambiar dependiendo del formato
    public void refreshGrid() throws Exception
    {

        String valores[] = dibujoService.getProgressInfo(rev);
        double percentage;

        progressLabel.setValue(valores[0] + "/" + valores[1] + " : "
                + valores[2]);

        funcLabel.setValue(valores[3] + "/" + valores[4]);

        if (valores[1].equals("?"))
            percentage = 0;
        else
            percentage = (Double.parseDouble(valores[0]) / Double.parseDouble(valores[1])) * 100;

        postes.setValue((int) percentage);

        if (valores[4].equals("?"))
            percentage = 0;
        else
            percentage = (Double.parseDouble(valores[3].replace(',', '.')) / Double.parseDouble(valores[4].replace(
                    ',', '.'))) * 100;

        function.setValue((int) percentage);

        rev = dibujoService.getRevision(ver, numRevision);
        if (rev.getCalculated())
        {
            timer.stop();
            Messagebox.show("Revisión " + numRevision + " de la versión "
                    + numVersion + " completada.", "Información",
                    Messagebox.OK, Messagebox.INFORMATION,
                    new org.zkoss.zk.ui.event.EventListener<Event>()
                    {
                        @Override
                        public void onEvent(Event e) throws Exception
                        {

                            if (e.getName().equals("onOK"))
                            {
                                // Redirect back
                                Executions.getCurrent().sendRedirect(
                                        "/drawing/");
                            }

                        }
                    });
        }

    }

    @Listen("onTimer = #timer")
    public void timer() throws Exception
    {
        refreshGrid();
    }

    @Listen("onClick = #goBack")
    public void doBackClick()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/drawing/");

    }

}

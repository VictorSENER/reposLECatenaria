/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.text.SimpleDateFormat;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Label;
import org.zkoss.zul.Progressmeter;
import org.zkoss.zul.Timer;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ReplanteoProgressPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    int numVersion;

    int numRevision;

    // Dialog components
    @Wire
    Progressmeter postes;
    @Wire
    Label progressLabel;
    @Wire
    Label version;
    @Wire
    Label revision;
    @Wire
    Label fecha;
    @Wire
    Timer timer;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        numVersion = (Integer) Executions.getCurrent().getAttribute(
                "numVersion");
        numRevision = (Integer) Executions.getCurrent().getAttribute(
                "numRevision");

        refreshGrid();

    }

    public void refreshGrid() throws Exception
    {
        Project project = projectService.getProjectById(actProj.getIdActive(session));

        ReplanteoVersion ver = replanteoService.getVersion(project, numVersion);
        ReplanteoRevision rev = replanteoService.getRevision(ver, numRevision);

        version.setValue("Version: " + numVersion);
        revision.setValue("Revision: " + numRevision);
        fecha.setValue("Fecha: "
                + new SimpleDateFormat("dd-MM-yyyy").format(rev.getDate()));

        String valores[] = replanteoService.getProgressInfo(rev);
        double percentage;

        progressLabel.setValue(valores[0] + "/" + valores[1]);

        if (valores[1].equals("?"))
            percentage = 0;
        else
            percentage = (Double.parseDouble(valores[0]) / Double.parseDouble(valores[1])) * 100;

        postes.setValue((int) percentage);
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
        Executions.getCurrent().sendRedirect("/replanteo");

    }

}

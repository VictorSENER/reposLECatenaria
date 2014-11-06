/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.ArrayList;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Cell;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Image;
import org.zkoss.zul.Label;
import org.zkoss.zul.Row;
import org.zkoss.zul.Rows;

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.PendoladoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class PendoladoErrorPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    int numVersion;

    int numRevision;

    // Dialog components
    @Wire
    Grid errorList;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        numVersion = (Integer) Executions.getCurrent().getAttribute(
                "numVersion");
        numRevision = (Integer) Executions.getCurrent().getAttribute(
                "numRevision");
        Project project = projectService.getProjectById(actProj.getIdActive(session));
        PendoladoVersion version = pendoladoService.getVersion(project,
                numVersion);
        PendoladoRevision revision = pendoladoService.getRevision(version,
                numRevision);

        ArrayList<String[]> errorLog = pendoladoService.getErrorLog(revision);

        String path = "/img/";

        Rows rows = new Rows();
        rows.setParent(errorList);

        for (int i = 0; i < errorLog.size(); i++)
        {

            Row row = new Row();
            Label idLabel;

            Cell cell0 = new Cell();
            if (errorLog.get(i)[0].equals("Error"))
            {
                Image imgError = new Image(path + "fatalerror.png");
                imgError.setWidth("15px");
                imgError.setHeight("15px");
                imgError.setParent(cell0);
            }

            else
            {
                Image imgWar = new Image(path + "warning.png");
                imgWar.setWidth("12px");
                imgWar.setHeight("12px");
                imgWar.setParent(cell0);
            }

            Cell cell1 = new Cell();
            idLabel = new Label(errorLog.get(i)[0]);
            idLabel.setParent(cell1);

            Cell cell2 = new Cell();
            idLabel = new Label(errorLog.get(i)[1]);
            idLabel.setParent(cell2);

            cell0.setParent(row);
            cell1.setParent(row);
            cell2.setParent(row);

            row.setParent(rows);

        }

    }

    @Listen("onClick = #goBack")
    public void doBackClick()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/pendolado");

    }

}

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
import org.zkoss.zul.Label;
import org.zkoss.zul.Row;
import org.zkoss.zul.Rows;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ReplanteoErrorPage extends SelectorComposer<Component>
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
        Project project = projectService.getProjectById(actProj.getIdActive(session));
        ReplanteoVersion version = replanteoService.getVersion(project,
                numVersion);
        ReplanteoRevision revision = replanteoService.getRevision(version,
                numRevision);

        ArrayList<String> errorLog = replanteoService.getErrorLog(revision);

        Rows rows = new Rows();
        rows.setParent(errorList);

        for (int i = 0; i < errorLog.size(); i += 2)
        {

            Row row = new Row();
            Label idLabel;

            Cell cell1 = new Cell();
            idLabel = new Label(errorLog.get(i));
            idLabel.setParent(cell1);

            Cell cell2 = new Cell();
            idLabel = new Label(errorLog.get(i + 1));
            idLabel.setParent(cell2);

            cell1.setParent(row);
            cell2.setParent(row);
            row.setParent(rows);

        }

    }

    @Listen("onClick = #goBack")
    public void doBackClick()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/replanteo");

    }

}

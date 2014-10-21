/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.SerializableEventListener;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Image;
import org.zkoss.zul.Label;
import org.zkoss.zul.Row;
import org.zkoss.zul.Rows;

import com.sener.sireca.web.service.SidebarLink;
import com.sener.sireca.web.service.SidebarLinksService;
import com.sener.sireca.web.service.SidebarLinksServiceImpl;

public class SidebarComponent extends SelectorComposer<Component>
{

    private static final long serialVersionUID = 1L;
    @Wire
    Grid fnList;

    // wire service
    SidebarLinksService linksService = new SidebarLinksServiceImpl();

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        // to initial view after view constructed.
        Rows rows = fnList.getRows();

        for (SidebarLink link : linksService.getLinks())
        {
            Row row = constructSidebarRow(link.getName(), link.getLabel(),
                    link.getIconUri(), link.getUrl());
            rows.appendChild(row);
        }
    }

    private Row constructSidebarRow(String name, String label, String imageSrc,
            final String locationUrl)
    {

        // construct component and hierarchy
        Row row = new Row();
        Image image = new Image(imageSrc);
        image.setWidth("24px");
        image.setHeight("24px");
        Label lab = new Label(label);

        row.appendChild(image);
        row.appendChild(lab);

        // set style attribute
        row.setSclass("sidebar-fn");

        EventListener<Event> actionListener = new SerializableEventListener<Event>()
        {
            private static final long serialVersionUID = 1L;

            @Override
            public void onEvent(Event event) throws Exception
            {
                // redirect current url to new location
                Executions.getCurrent().sendRedirect(locationUrl);
            }
        };

        row.addEventListener(Events.ON_CLICK, actionListener);

        return row;
    }
}

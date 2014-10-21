/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;

public class BannerComponent extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    @Listen("onClick=#logout")
    public void doLogout()
    {
        Executions.sendRedirect("logout");
    }
}

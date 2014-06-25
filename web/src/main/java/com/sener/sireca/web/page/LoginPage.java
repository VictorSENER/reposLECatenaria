/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */
package com.sener.sireca.web.page;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Label;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.service.AuthenticationService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class LoginPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    @Wire
    Textbox username;
    @Wire
    Textbox password;
    @Wire
    Label message;

    @Listen("onClick=#login; onOK=#loginWin")
    public void doLogin()
    {
        // Validate user authentication
        HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();
        String username = this.username.getValue();
        String password = this.password.getValue();
        AuthenticationService authService = (AuthenticationService) SpringApplicationContext.getBean("authService");
        if (!authService.login(session, username, password))
        {
            message.setValue("Credenciales no validas.");
            return;
        }

        // Valid authentication
        // Show main page.
        Executions.sendRedirect("/main");
    }
}

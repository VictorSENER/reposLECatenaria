/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.List;

import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.service.UserService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class UserPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button addUser;
    @Wire
    Listbox userListbox;
    @Wire
    Component selectedUserBlock;
    @Wire
    Textbox selectedUserUsername;
    @Wire
    Textbox selectedUserPassword;
    @Wire
    Button updateSelectedUser;

    // Users list
    ListModelList<User> userListModel;

    // User currently selected.
    User selectedUser;

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        // Fill users list using DB data
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        List<User> userList = userService.getAllUsers();
        userListModel = new ListModelList<User>(userList);
        userListbox.setModel(userListModel);
    }

    @Listen("onClick = #addUser")
    public void doUserAdd()
    {
        // Get a username for the new user.
        String username = buildNewUsername();

        // Create new user object.
        User user = new User();
        user.setUsername(username);
        user.setPassword("");

        // Store new user into DB.
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        userService.insertUser(user);

        // Add new user into list model and select it.
        selectedUser = userService.getUserByUsername(username);
        userListModel.add(selectedUser);
        userListModel.addToSelection(selectedUser);

        // Refresh detail view: new user will be shown.
        refreshDetailView();
    }

    @Listen("onUserDelete = #userListbox")
    public void doUserDelete(final ForwardEvent evt)
    {
        // Ask for user confirmation.
        Messagebox.show("Está seguro que quiere borrar a este usuario?",
                "Confirmación", Messagebox.OK | Messagebox.CANCEL,
                Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    @Override
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                        {
                            // Get user to be deleted.
                            Button btn = (Button) evt.getOrigin().getTarget();
                            Listitem litem = (Listitem) btn.getParent().getParent();
                            User user = (User) litem.getValue();

                            // Delete user from DB.
                            UserService userService = (UserService) SpringApplicationContext.getBean("userService");
                            userService.deleteUser(user.getId());

                            // Remove user from listbox.
                            userListModel.remove(user);

                            // Refresh view when necessary.
                            if (user.equals(selectedUser))
                            {
                                selectedUser = null;
                                refreshDetailView();
                            }

                            // Show confirmation.
                            Clients.showNotification("Usuario borrado correctamente");
                        }
                    }
                });
    }

    @Listen("onSelect = #userListbox")
    public void doUserSelect()
    {
        // Update selected user member
        if (userListModel.isSelectionEmpty())
            selectedUser = null;
        else
            selectedUser = userListModel.getSelection().iterator().next();

        // Refresh view according to new selection.
        refreshDetailView();
    }

    @Listen("onClick = #cancelSelectedUser")
    public void doCancelClick()
    {

        Messagebox.show("Está seguro que quiere cancelar?", "Confirmación",
                Messagebox.OK | Messagebox.CANCEL, Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    @Override
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                        {
                            selectedUser = null;

                            // Refresh view according to new selection.
                            refreshDetailView();
                        }
                    }
                });

    }

    @Listen("onClick = #updateSelectedUser")
    public void doUpdateClick()
    {
        // Checks if username is empty.
        if (Strings.isBlank(selectedUserUsername.getValue()))
        {
            Clients.showNotification(
                    "El nombre de usuario no puede estar vacío.",
                    selectedUserUsername);
            return;
        }

        else if (selectedUserUsername.getValue().length() > 50)
        {
            Clients.showNotification(
                    "El nombre de usuario no puede ser tan largo. (Máximo 50 carácteres)",
                    selectedUserUsername);
            return;
        }

        // Checks if password is empty.
        if (Strings.isBlank(selectedUserPassword.getValue()))
        {
            Clients.showNotification("El password no puede estar vacío.",
                    selectedUserPassword);
            return;
        }

        else if (selectedUserPassword.getValue().length() > 50)
        {
            Clients.showNotification(
                    "El password no puede ser tan largo. (Máximo 50 carácteres)",
                    selectedUserPassword);
            return;
        }
        // Set new data to selected user.
        selectedUser.setUsername(selectedUserUsername.getValue());
        selectedUser.setPassword(selectedUserPassword.getValue());

        // Save new data into DB.
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        userService.updateUser(selectedUser);

        // Update user into listbox.
        userListModel.set(userListModel.indexOf(selectedUser), selectedUser);

        // Show message for user.
        Clients.showNotification("Usuario guardado correctamente");

        selectedUser = null;

        // Refresh view according to new selection.
        refreshDetailView();

    }

    private void refreshDetailView()
    {
        // Check if there's a user selected.
        if (selectedUser == null)
        {
            // No user selected.
            selectedUserBlock.setVisible(false);
            selectedUserUsername.setValue(null);
            selectedUserPassword.setValue(null);
            updateSelectedUser.setDisabled(true);
        }
        else
        {
            // User selected.
            selectedUserBlock.setVisible(true);
            selectedUserUsername.setValue(selectedUser.getUsername());
            selectedUserPassword.setValue(selectedUser.getPassword());
            updateSelectedUser.setDisabled(false);
        }
    }

    private String buildNewUsername()
    {
        // Check if base username isn't used.
        String baseUsername = "nuevo";
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");
        if (userService.getUserByUsername(baseUsername) == null)
            return baseUsername;

        int sequential = 1;
        while (sequential < 100)
        {
            String seqUsername = baseUsername + sequential;
            if (userService.getUserByUsername(seqUsername) == null)
                return seqUsername;

            sequential++;
        }

        return "nuevo";
    }
}

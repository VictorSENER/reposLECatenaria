/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.worker;

import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.service.ReplanteoServiceImpl;

public class ReplanteoWorker extends Thread
{
    // Revisión de la cual calcular el cuaderno de replanteo
    private ReplanteoRevision revision;

    public ReplanteoWorker(ReplanteoRevision revision)
    {
        this.revision = revision;
    }

    @Override
    public void run()
    {
        ReplanteoServiceImpl service = new ReplanteoServiceImpl();
        service.calculateRevision(this.revision);
    }
}

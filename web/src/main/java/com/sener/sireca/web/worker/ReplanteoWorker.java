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
    private int idCatenaria;
    private int pkIni;
    private int pkFin;

    public ReplanteoWorker(ReplanteoRevision revision, int idCatenaria,
            int pkIni, int pkFin)
    {
        super();
        this.revision = revision;
        this.idCatenaria = idCatenaria;
        this.pkIni = pkIni;
        this.pkFin = pkFin;
    }

    @Override
    public void run()
    {
        ReplanteoServiceImpl service = new ReplanteoServiceImpl();
        service.calculateRevision(this.revision, this.idCatenaria, this.pkIni,
                this.pkFin);
    }
}

\\172.16.79.28\Ninstaladores\SQ_DOFIS\ZSQ_DOFIS.exe "\\172.16.79.28\Ninstaladores\SQ_DOFIS\Activos.txt" "SELECT P1.CO_TRAB STCD1, P1.TI_SITU STATX, N1.NO_APEL_PATE + # # + N1.NO_APEL_MATE + #, # + N1.NO_TRAB NAME1, C1.CO_CENT_COST KOSTL FROM TMTRAB_EMPR P1 INNER JOIN TMTRAB_PERS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN TDCCOS_TRAB C1 ON P1.CO_TRAB = C1.CO_TRAB WHERE P1.CO_EMPR = #10# AND P1.TI_SITU = #ACT# AND SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR ),1,3) = #011#"
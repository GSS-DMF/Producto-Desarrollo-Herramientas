// Funcion que crea un SAPModel a partir de un SAPObject.
// Te devuelve el SAPModel para poder usarlo en otras funciones.
// Se necesita como input un SAPObject.

public cSapModel OpenSAPModel(cOAPI SapObject)
{
    mySapModel = SapObject.SapModel;
    mySapModel.InitializeNewModel();

    return mySapModel;
}
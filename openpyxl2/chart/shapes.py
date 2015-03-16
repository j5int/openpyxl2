from openpyxl2.descriptors.serialisable import Serialisable


from openpyxl2.descriptors import (
    Set,
    )

class Shape(Serialisable):

    val = Set(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


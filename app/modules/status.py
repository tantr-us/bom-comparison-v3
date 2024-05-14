from enum import IntEnum

# Bom Comparison status code
class BcStatus(IntEnum):
    PART_NUMBER_MISMATCH = 10
    DESCRIPTION_MISMATCH = 20
    UOM_MISMATCH = 30
    QTY_CHANGE = 40
    REV_CHANGE = 50
    REF_DES_MISMATCH = 60
    PART_ADD = 70
    PART_REMOVE = 80
    MFR_NAME_MISMATCH = 90
    MFR_NUMBER_MISMATCH = 100
    AVL_ADD = 110
    AVL_REMOVE = 120
    AVL_MISMATCH = 130 # use to hightligth the mismatch row

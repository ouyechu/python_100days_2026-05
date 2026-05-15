PDU_ASSET_STATUS_ENUM = [
    "使用中-已上电已上负载",
    "使用中-已上电未上负载",
    "停用中",
    "故障中",
    "维修中",
    "已报废",
]

GENERAL_ASSET_STATUS_ENUM = ["使用中", "停用中", "故障中", "维修中", "已报废"]

# 后端统一校验集合（兼容历史值 + PDU 细分值）
ASSET_STATUS_ENUM = ["使用中"] + [s for s in PDU_ASSET_STATUS_ENUM if s != "使用中"]

IT_RACK_MOUNT_STATUS_ENUM = ["已上架", "停用", "未上架"]


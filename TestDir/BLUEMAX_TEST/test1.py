"""경고 페이지 설정 REST API"""
from app.core.decorator import permission_required
from app.core.roles import Permission
from app.core.viewsets import GenericViewSet
from app.rest.resource.sm.manager_user_dao import ManagerUserListDao, ManagerUserApplyDao, ManagerUserCancelDao, \
    ManagerBookmarkDao, ManagerBookmarkApplyDao, ManagerBookmarkCancelDao, ManagerUserClientDao


class ManagerUserList(GenericViewSet):
    dao_class = ManagerUserListDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_user_list_class = ManagerUserList.create_view({
    'get': 'list',
    'delete': 'multi_destroy',
    'post': 'create',
})

manager_user_class = ManagerUserList.create_view({
    'get': 'retrieve',
    'put': 'update',
    'delete': 'destroy',
})


class ManagerUserClient(GenericViewSet):
    dao_class = ManagerUserClientDao

    @permission_required(Permission.READ)
    def create(self, *args, **kwargs):
        return super().create(**kwargs)

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_user_client_gen_class = ManagerUserClient.create_view({
    'get': 'retrieve',
    'post': 'create',
    'put': 'update'
})

class ManagerUserApply(GenericViewSet):
    dao_class = ManagerUserApplyDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_user_apply_class = ManagerUserApply.create_view({
    'put': 'update',
})


class ManagerUserCancel(GenericViewSet):
    dao_class = ManagerUserCancelDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_user_cancel_class = ManagerUserCancel.create_view({
    'put': 'update',
})


class ManagerBookmark(GenericViewSet):
    dao_class = ManagerBookmarkDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_bookmark_class = ManagerBookmark.create_view({
    'get': 'retrieve',
    'put': 'update'
})


class ManagerBookmarkApply(GenericViewSet):
    dao_class = ManagerBookmarkApplyDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_bookmark_apply_class = ManagerBookmarkApply.create_view({
    'put': 'update',
})


class ManagerBookmarkCancel(GenericViewSet):
    dao_class = ManagerBookmarkCancelDao

    @permission_required(Permission.READ)
    def update(self, *args, **kwargs):
        return super().update(**kwargs)


manager_bookmark_cancel_class = ManagerBookmarkCancel.create_view({
    'put': 'update',
})

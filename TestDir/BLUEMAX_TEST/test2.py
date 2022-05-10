"""경고 페이지 설정 DAO"""
import time
import uuid
import logging
import hashlib

from flask import current_app

from app.common.api import LcappsConfigSimpleDao, LcappsConfigDao
from app.common.const import HTTP_ERROR_CODES
from app.common.exceptions import RestApiAuthError
from app.data.base_system import get_vs_mode
from app.data.menu import MenuConfig
from app.rest.main.daos import ListConfigDao
from app.utils.security import AESCipher

logger = logging.getLogger(__name__)


class ManagerUserListDao(ListConfigDao):
    pk_names = "adm_id"

    list_procedure = "GET_ADM"
    retrieve_procedure = "GET_ADM"
    create_procedure = "SET_ADM"
    update_procedure = "MOD_ADM"
    destroy_procedure = "DEL_ADM"

    def __init__(self, **kwargs):
        self.vs_mode = get_vs_mode(current_app.config['VS-INFO'])
        super().__init__(**kwargs)

    def list(self, *args, **kwargs):
        if self.vs_mode == MenuConfig.HOST_MENU:
            self.list_procedure = 'GET_MGMT_ADM'

        return super().list(args, **kwargs)

    def retrieve(self, pk, **kwargs):
        if self.vs_mode == MenuConfig.HOST_MENU:
            self.retrieve_procedure = 'GET_MGMT_ADM'

        access_level = int(self.session_data.get('access_level', 1))

        user_dao_data = super().retrieve(pk)
        org_level = user_dao_data.get('level', 1)

        isValid = False
        if access_level == 1 and int(self.user_key) == pk:
            isValid = True
        elif access_level > 1:
            if access_level >= int(org_level):
                isValid = True

        if not isValid:
            raise RestApiAuthError(94017, HTTP_ERROR_CODES[94017])

        result = super().retrieve(pk, **kwargs)

        return result

    def create(self, **kwargs):
        if self.vs_mode == MenuConfig.HOST_MENU:
            self.create_procedure = 'SET_MGMT_ADM'

        access_level = int(self.session_data.get('access_level', 1))
        level = int(kwargs.get('level', 1))

        if level > access_level:
            raise RestApiAuthError(94017, HTTP_ERROR_CODES[94017])

        try:
            if kwargs.get('login_pw', '') and not current_app.config.get('DEBUG'):
                cipher = AESCipher(self.session_data.get('aes_key'))
                login_pw = cipher.decrypt_and_b64encode(kwargs.get('login_pw'))
                kwargs.update({'login_pw': login_pw})
        except Exception as e:
            logger.error(e)

        return super().create(**kwargs)

    def update(self, pk, **kwargs):
        if self.vs_mode == MenuConfig.HOST_MENU:
            self.update_procedure = 'MOD_MGMT_ADM'

        access_level = int(self.session_data.get('access_level', 1))
        level = int(kwargs.get('level', 1))

        isValid = False
        if access_level == 1 and int(self.user_key) == pk:
            isValid = True
        elif access_level > 1:
            if access_level >= level:
                isValid = True

        if not isValid:
            raise RestApiAuthError(94017, HTTP_ERROR_CODES[94017])

        try:
            if kwargs.get('login_pw', '') and not current_app.config.get('DEBUG'):
                cipher = AESCipher(self.session_data.get('aes_key'))
                login_pw = cipher.decrypt_and_b64encode(kwargs.get('login_pw'))
                kwargs.update({'login_pw': login_pw})
        except Exception as e:
            logger.error(e)

        return super().update(pk, **kwargs)

    def destroy(self, pk, **kwargs):
        if self.vs_mode == MenuConfig.HOST_MENU:
            self.destroy_procedure = 'DEL_MGMT_ADM'

        access_level = int(self.session_data.get('access_level', 1))
        user_dao_data = super().retrieve(pk)
        org_level = user_dao_data.get('level', 1)

        if org_level > access_level:
            raise RestApiAuthError(94017, HTTP_ERROR_CODES[94017])

        return super().destroy(pk, **kwargs)

    def init_destroy(self, pk, **kwargs):

        return self.call_api('DEL_ADM_IGNORE_SESS', **kwargs)


class ManagerUserClientDao(LcappsConfigDao):
    retrieve_procedure = 'GET_ADM_EXT_CLNT'

    def retrieve(self, pk, **kwargs):
        access_level = int(self.session_data.get('access_level', 1))

        dao_kwargs = self.get_gui_params()
        dao_kwargs.update({'access_level': str(access_level)})
        user_dao = ManagerUserListDao(**dao_kwargs)
        user_dao_data = user_dao.retrieve(pk, **{})
        org_level = user_dao_data.get('level', 1)

        isValid = False
        if access_level == 1 and int(self.user_key) == pk:
            isValid = True
        elif access_level > 1:
            if access_level >= int(org_level):
                isValid = True

        if not isValid:
            raise RestApiAuthError(94017, HTTP_ERROR_CODES[94017])

        kwargs.update({'adm_id': pk})
        return self.call_api(self.retrieve_procedure, **kwargs)

    def create(self, pk, **kwargs):
        access_level = int(self.session_data.get('access_level', 1))

        cl_id = uuid.uuid4().hex
        secret_key = current_app.config['SECRET_KEY']
        body = {
            'ext_clnt_id': cl_id,
            'ext_clnt_secret': hashlib.sha256(secret_key.encode() + str(time.time()).encode()).hexdigest()
        }

        dao_kwargs = self.get_gui_params()
        dao_kwargs.update({'access_level': str(access_level)})
        user_dao = ManagerUserListDao(**dao_kwargs)
        user_dao.update(pk, **body)

        return body

    def update(self, pk, **kwargs):
        access_level = int(self.session_data.get('access_level', 1))

        secret_key = current_app.config['SECRET_KEY']
        body = {
            'ext_clnt_secret': hashlib.sha256(secret_key.encode() + str(time.time()).encode()).hexdigest()
        }
        dao_kwargs = self.get_gui_params()
        dao_kwargs.update({'access_level': str(access_level)})
        user_dao = ManagerUserListDao(**dao_kwargs)
        user_dao.update(pk, **body)

        return body

class ManagerUserApplyDao(LcappsConfigSimpleDao):
    update_procedure = "ADM_COMMIT"


class ManagerUserCancelDao(LcappsConfigSimpleDao):
    update_procedure = "ADM_ROLLBACK"


class ManagerBookmarkDao(ListConfigDao):
    pk_names = 'adm_id'
    retrieve_procedure = 'GET_ADM_BKMK'
    update_procedure = 'SET_ADM_BKMK'


class ManagerBookmarkApplyDao(LcappsConfigSimpleDao):
    update_procedure = "ADM_BKMK_COMMIT"


class ManagerBookmarkCancelDao(LcappsConfigSimpleDao):
    update_procedure = "ADM_BKMK_ROLLBACK"

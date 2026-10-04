"""
存储抽象层
提供统一的 JSON 和 Blob 存储接口，支持多种后端

支持的存储模式：
  - local: 本地文件系统（开发/测试用）
  - cloud: 云端存储（通过环境变量配置后端）

Cloud 模式后端可插拔：
  - KV 后端: env（环境变量） / makers（Makers 原生 KV） / cos（腾讯云COS）
  - Blob 后端: cos（腾讯云COS） / makers（Makers 原生 Blob）
"""

import os
import json
import base64
from abc import ABC, abstractmethod


# ==================== KV 存储后端 ====================

class KVBackend(ABC):
    """KV 存储后端基类"""

    @abstractmethod
    def get(self, key):
        """获取值，返回字符串或 None"""
        pass

    @abstractmethod
    def put(self, key, value):
        """存储值（字符串）"""
        pass

    @abstractmethod
    def delete(self, key):
        """删除键"""
        pass

    @abstractmethod
    def list_keys(self, prefix=''):
        """列出所有键"""
        pass


class EnvKVBackend(KVBackend):
    """环境变量 KV 后端 - 适合存储少量配置数据

    数据以 JSON 格式存在单个环境变量中，键名作为 JSON 的 key
    环境变量名：KV_STORE_DATA
    """

    ENV_VAR = 'KV_STORE_DATA'

    def __init__(self):
        self._data = self._load()

    def _load(self):
        raw = os.environ.get(self.ENV_VAR, '{}')
        try:
            return json.loads(raw)
        except (json.JSONDecodeError, TypeError):
            return {}

    def get(self, key):
        return self._data.get(key)

    def put(self, key, value):
        self._data[key] = value
        # 注意：环境变量方式下，put 只更新内存，不持久化
        # 如果需要持久化，需要使用 COS 或 Makers KV

    def delete(self, key):
        if key in self._data:
            del self._data[key]

    def list_keys(self, prefix=''):
        return [k for k in self._data.keys() if k.startswith(prefix)]


class MakersKVBackend(KVBackend):
    """Makers 原生 KV 后端

    部署到 EdgeOne Makers 后使用，通过全局 KV 对象访问。
    具体 API 以 Makers Python 运行时实际提供的为准。
    """

    def __init__(self, binding_name='GRADING_KV'):
        self.binding_name = binding_name
        self._kv = None
        # 尝试获取 Makers KV 绑定
        # 注意：实际 API 以 Makers Python 运行时为准
        try:
            # Makers 可能通过全局变量注入 KV 实例
            import __main__
            self._kv = getattr(__main__, binding_name, None)
        except Exception:
            pass

        if self._kv is None:
            # 回退：尝试从环境变量获取
            fallback = os.environ.get(binding_name)
            if fallback:
                try:
                    self._kv = json.loads(fallback)
                except Exception:
                    self._kv = {}
            else:
                self._kv = {}

    def get(self, key):
        if hasattr(self._kv, 'get'):
            # 如果是 dict 或支持 get 方法的对象
            return self._kv.get(key)
        return None

    def put(self, key, value):
        if hasattr(self._kv, '__setitem__'):
            self._kv[key] = value
        # 如果是 Makers 原生 KV 对象，可能有 put 方法
        elif hasattr(self._kv, 'put'):
            self._kv.put(key, value)

    def delete(self, key):
        if hasattr(self._kv, '__delitem__'):
            if key in self._kv:
                del self._kv[key]
        elif hasattr(self._kv, 'delete'):
            self._kv.delete(key)

    def list_keys(self, prefix=''):
        if hasattr(self._kv, 'keys'):
            return [k for k in self._kv.keys() if k.startswith(prefix)]
        return []


# ==================== Blob 存储后端 ====================

class BlobBackend(ABC):
    """Blob 存储后端基类"""

    @abstractmethod
    def get(self, key):
        """获取二进制数据，返回 bytes 或 None"""
        pass

    @abstractmethod
    def put(self, key, data_bytes):
        """存储二进制数据"""
        pass

    @abstractmethod
    def delete(self, key):
        """删除对象"""
        pass

    @abstractmethod
    def list(self, prefix=''):
        """列出对象键名"""
        pass


class LocalBlobBackend(BlobBackend):
    """本地文件系统 Blob 后端"""

    def __init__(self, base_dir='templates_user/blobs'):
        self.base_dir = base_dir
        os.makedirs(base_dir, exist_ok=True)

    def _path(self, key):
        # 防止路径穿越
        safe_key = key.replace('..', '').lstrip('/')
        return os.path.join(self.base_dir, safe_key)

    def get(self, key):
        path = self._path(key)
        if os.path.exists(path):
            with open(path, 'rb') as f:
                return f.read()
        return None

    def put(self, key, data_bytes):
        path = self._path(key)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'wb') as f:
            f.write(data_bytes)

    def delete(self, key):
        path = self._path(key)
        if os.path.exists(path):
            os.remove(path)
            return True
        return False

    def list(self, prefix=''):
        blobs = []
        if os.path.exists(self.base_dir):
            for root, dirs, files in os.walk(self.base_dir):
                for fname in files:
                    full_path = os.path.join(root, fname)
                    rel_path = os.path.relpath(full_path, self.base_dir)
                    if rel_path.startswith(prefix):
                        blobs.append(rel_path)
        return blobs


class COSBlobBackend(BlobBackend):
    """腾讯云 COS Blob 后端

    需要安装 cos-python-sdk-v5
    环境变量：
      - COS_SECRET_ID
      - COS_SECRET_KEY
      - COS_REGION
      - COS_BUCKET
    """

    def __init__(self):
        secret_id = os.environ.get('COS_SECRET_ID', '')
        secret_key = os.environ.get('COS_SECRET_KEY', '')
        region = os.environ.get('COS_REGION', '')
        bucket = os.environ.get('COS_BUCKET', '')

        self.bucket = bucket
        self.client = None

        if secret_id and secret_key and region and bucket:
            try:
                from qcloud_cos import CosConfig, CosS3Client
                config = CosConfig(Region=region, SecretId=secret_id, SecretKey=secret_key)
                self.client = CosS3Client(config)
            except ImportError:
                print('[Storage] cos-python-sdk-v5 未安装，COS Blob 后端不可用')
            except Exception as e:
                print(f'[Storage] COS 初始化失败: {e}')

    def _check(self):
        if self.client is None:
            raise RuntimeError('COS 客户端未初始化，请配置 COS_SECRET_ID/COS_SECRET_KEY/COS_REGION/COS_BUCKET')

    def get(self, key):
        self._check()
        try:
            response = self.client.get_object(Bucket=self.bucket, Key=key)
            return response['Body'].get_raw_content()
        except Exception:
            return None

    def put(self, key, data_bytes):
        self._check()
        from io import BytesIO
        self.client.put_object(
            Bucket=self.bucket,
            Key=key,
            Body=BytesIO(data_bytes)
        )

    def delete(self, key):
        self._check()
        try:
            self.client.delete_object(Bucket=self.bucket, Key=key)
            return True
        except Exception:
            return False

    def list(self, prefix=''):
        self._check()
        try:
            response = self.client.list_objects(Bucket=self.bucket, Prefix=prefix)
            return [item['Key'] for item in response.get('Contents', [])]
        except Exception:
            return []


class MakersBlobBackend(BlobBackend):
    """Makers 原生 Blob 后端

    部署到 EdgeOne Makers 后使用。
    具体 API 以 Makers Python 运行时实际提供的为准。
    """

    def __init__(self, binding_name='GRADING_BLOB'):
        self.binding_name = binding_name
        self._store = None
        try:
            import __main__
            self._store = getattr(__main__, binding_name, None)
        except Exception:
            pass

    def _check(self):
        if self._store is None:
            raise RuntimeError(f'Makers Blob 绑定 {self.binding_name} 不可用')

    def get(self, key):
        self._check()
        if hasattr(self._store, 'get'):
            return self._store.get(key)
        elif hasattr(self._store, 'get_object'):
            return self._store.get_object(key)
        return None

    def put(self, key, data_bytes):
        self._check()
        if hasattr(self._store, 'set'):
            self._store.set(key, data_bytes)
        elif hasattr(self._store, 'put'):
            self._store.put(key, data_bytes)
        elif hasattr(self._store, 'put_object'):
            self._store.put_object(key, data_bytes)

    def delete(self, key):
        self._check()
        if hasattr(self._store, 'delete'):
            return self._store.delete(key)
        elif hasattr(self._store, 'delete_object'):
            return self._store.delete_object(key)
        return False

    def list(self, prefix=''):
        self._check()
        if hasattr(self._store, 'list'):
            return self._store.list(prefix=prefix)
        elif hasattr(self._store, 'list_objects'):
            return self._store.list_objects(prefix=prefix)
        return []


# ==================== 统一 Storage 类 ====================

class Storage:
    """统一存储抽象类

    提供 JSON 和 Blob 两种存储方式的统一接口。
    根据 mode 参数选择不同的后端实现。
    """

    def __init__(self, mode='local', base_dir='templates_user',
                 kv_backend=None, blob_backend=None):
        """
        初始化存储
        :param mode: 存储模式 'local' | 'cloud'
        :param base_dir: 本地存储基础目录（local 模式用）
        :param kv_backend: 自定义 KV 后端（cloud 模式用，为 None 则自动选择）
        :param blob_backend: 自定义 Blob 后端（cloud 模式用，为 None 则自动选择）
        """
        self.mode = mode
        self.base_dir = base_dir

        if mode == 'local':
            self._kv = None  # JSON 文件方式，直接操作文件
            self._blob = LocalBlobBackend(os.path.join(base_dir, 'blobs'))
            self._ensure_local_dirs()
        elif mode == 'cloud':
            self._kv = kv_backend or self._auto_kv_backend()
            self._blob = blob_backend or self._auto_blob_backend()
        else:
            raise ValueError(f'不支持的存储模式: {mode}')

    def _ensure_local_dirs(self):
        os.makedirs(self.base_dir, exist_ok=True)

    def _auto_kv_backend(self):
        """自动选择 KV 后端"""
        # 优先尝试 Makers 原生 KV
        makers_kv_name = os.environ.get('KV_BINDING_NAME', 'GRADING_KV')
        kv_backend = MakersKVBackend(makers_kv_name)
        # 如果 Makers KV 可用（有数据或对象有 put 方法），使用它
        if hasattr(kv_backend._kv, 'put') or len(kv_backend.list_keys()) > 0:
            return kv_backend
        # 回退到环境变量
        return EnvKVBackend()

    def _auto_blob_backend(self):
        """自动选择 Blob 后端"""
        # 优先尝试 Makers 原生 Blob
        makers_blob_name = os.environ.get('BLOB_BINDING_NAME', 'GRADING_BLOB')
        blob_backend = MakersBlobBackend(makers_blob_name)
        if blob_backend._store is not None:
            return blob_backend
        # 尝试 COS
        if os.environ.get('COS_BUCKET'):
            cos_backend = COSBlobBackend()
            if cos_backend.client is not None:
                return cos_backend
        # 回退到本地（开发调试用）
        print('[Storage] Cloud 模式 Blob 后端不可用，回退到本地文件')
        return LocalBlobBackend(os.path.join(self.base_dir, 'blobs'))

    def _json_path(self, key):
        """获取 JSON 文件路径（local 模式）"""
        return os.path.join(self.base_dir, f'{key}.json')

    # ---------- JSON 操作 ----------

    def get_json(self, key):
        """
        读取 JSON 数据
        :param key: 键名
        :return: dict/list，如果不存在返回 None
        """
        if self.mode == 'local':
            path = self._json_path(key)
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        else:  # cloud
            raw = self._kv.get(key)
            if raw:
                try:
                    return json.loads(raw) if isinstance(raw, str) else raw
                except (json.JSONDecodeError, TypeError):
                    return None
            return None

    def put_json(self, key, value):
        """
        写入 JSON 数据
        :param key: 键名
        :param value: 要写入的数据（dict/list）
        """
        if self.mode == 'local':
            path = self._json_path(key)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(value, f, ensure_ascii=False, indent=2)
        else:  # cloud
            self._kv.put(key, json.dumps(value, ensure_ascii=False))

    def delete_json(self, key):
        """删除 JSON 数据"""
        if self.mode == 'local':
            path = self._json_path(key)
            if os.path.exists(path):
                os.remove(path)
                return True
            return False
        else:
            self._kv.delete(key)
            return True

    # ---------- Blob 操作 ----------

    def get_blob(self, key):
        """
        读取 Blob 数据（二进制）
        :param key: 键名
        :return: bytes，如果不存在返回 None
        """
        return self._blob.get(key)

    def put_blob(self, key, data_bytes):
        """
        写入 Blob 数据
        :param key: 键名
        :param data_bytes: bytes 数据
        """
        self._blob.put(key, data_bytes)

    def delete_blob(self, key):
        """
        删除 Blob
        :param key: 键名
        :return: 是否成功删除
        """
        return self._blob.delete(key)

    def list_blobs(self, prefix=''):
        """
        列出所有 Blob
        :param prefix: 前缀过滤
        :return: Blob 键名列表
        """
        return self._blob.list(prefix=prefix)

    # ---------- 本地文件操作（仅 local 模式） ----------

    def get_file_path(self, key):
        """
        获取本地文件的完整路径（仅 local 模式）
        :param key: 键名（文件名）
        :return: 完整路径
        """
        if self.mode == 'local':
            return os.path.join(self.base_dir, key)
        # cloud 模式下，下载到临时文件并返回路径
        data = self.get_blob(key)
        if data:
            import tempfile
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(key)[1])
            tmp.write(data)
            tmp.close()
            return tmp.name
        return None

    def file_exists(self, key):
        """
        检查文件是否存在
        :param key: 键名（文件名）
        :return: bool
        """
        if self.mode == 'local':
            path = os.path.join(self.base_dir, key)
            return os.path.exists(path)
        else:
            return self.get_blob(key) is not None


# ==================== 单例管理 ====================

_instance = None


def get_storage(mode=None, base_dir='templates_user'):
    """
    获取 Storage 单例
    :param mode: 存储模式，为 None 则从环境变量 STORAGE_MODE 读取，默认 local
    :param base_dir: 基础目录（local 模式）
    :return: Storage 实例
    """
    global _instance
    if _instance is None:
        if mode is None:
            mode = os.environ.get('STORAGE_MODE', 'local')
        _instance = Storage(mode=mode, base_dir=base_dir)
    return _instance


def reset_storage():
    """重置单例（用于测试或模式切换）"""
    global _instance
    _instance = None

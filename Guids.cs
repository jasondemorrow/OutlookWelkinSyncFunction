using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OutlookWelkinSyncFunction
{
    public static class Guids
    {
		private static readonly Guid NAMESPACE_ID = Guid.Parse("0b0bc2a4-a46a-4479-98b3-d66b58fe732b");

		/// <summary>
		/// Derives a consistent, deterministic GUID from some given text.
		/// </summary>
		/// <param name="text">The text from which to generate a consistent GUID.</param>
		/// <returns>A UUID derived from the text.</returns>
        public static Guid FromText(string text)
        {
			return FromNamespaceAndName(NAMESPACE_ID, text);
		}

		/// <summary>
		/// Creates a name-based UUID using the algorithm (version 3) from RFC 4122 §4.3.
		/// </summary>
		/// <param name="namespaceId">The ID of the namespace.</param>
		/// <param name="name">The name (within that namespace).</param>
		/// <returns>A UUID derived from the namespace and name.</returns>
        public static Guid FromNamespaceAndName(Guid namespaceId, string name)
        {
            var nameBytes = Encoding.UTF8.GetBytes(name);
			// convert the namespace UUID to network order
			var namespaceBytes = namespaceId.ToByteArray();
			SwapByteOrder(namespaceBytes);

			// compute the hash of the namespace ID concatenated with the name
			var data = namespaceBytes.Concat(nameBytes).ToArray();
			byte[] hash;
			using (HashAlgorithm algorithm = MD5.Create())
            {
				hash = algorithm.ComputeHash(data);
            }

			// most bytes from the hash are copied straight to the bytes of the new GUID
			var newGuid = new byte[16];
			Array.Copy(hash, 0, newGuid, 0, 16);

			// set the four most significant bits (bits 12 through 15) of the time_hi_and_version field to the appropriate 4-bit version number from Section 4.1.3
			newGuid[6] = (byte) ((newGuid[6] & 0x0F) | (3 << 4));

			// set the two most significant bits (bits 6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively (step 10)
			newGuid[8] = (byte) ((newGuid[8] & 0x3F) | 0x80);

			// convert the resulting UUID to local byte order (step 13)
			SwapByteOrder(newGuid);
			return new Guid(newGuid);
        }
        
		// Converts a GUID (expressed as a byte array) to/from network order (MSB-first).
		internal static void SwapByteOrder(byte[] guid)
		{
			SwapBytes(guid, 0, 3);
			SwapBytes(guid, 1, 2);
			SwapBytes(guid, 4, 5);
			SwapBytes(guid, 6, 7);
		}

		private static void SwapBytes(byte[] guid, int left, int right)
		{
			var temp = guid[left];
			guid[left] = guid[right];
			guid[right] = temp;
		}
    }
}